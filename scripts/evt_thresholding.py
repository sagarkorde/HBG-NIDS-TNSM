"""
HBG-NIDS EVT Thresholding Module
Implements Streaming Peak-Over-Threshold (SPOT) and Drift SPOT (DSPOT)
for adaptive, drift-robust NIDS anomaly detection.
"""

import numpy as np
from scipy.stats import genpareto

class SPOTDetector:
    """
    Streaming Peak-Over-Threshold (SPOT) adaptive thresholding detector.
    Uses Generalized Pareto Distribution (GPD) modeling of excesses above an initial threshold.
    """
    def __init__(self, probability=1e-4, q_percentile=0.98):
        """
        Parameters:
        - probability: Target probability of failure (P), e.g., 1e-4 or 0.05.
        - q_percentile: Quantile ratio (e.g., 0.98) defining the initial threshold q.
        """
        self.probability = probability
        self.q_percentile = q_percentile
        self.q = None
        self.c = None
        self.scale = None
        self.N = 0
        self.N_q = 0
        self.peaks = []
        self.baseline_scores = None

    def fit(self, baseline_scores):
        """
        Fits GPD to the initial baseline anomaly scores.
        """
        scores = np.asarray(baseline_scores)
        self.baseline_scores = scores
        self.N = len(scores)
        
        # Calculate initial threshold q
        self.q = np.percentile(scores, self.q_percentile * 100)
        
        # Extract peaks above q
        peaks = scores[scores > self.q] - self.q
        self.peaks = list(peaks)
        self.N_q = len(peaks)
        
        if self.N_q < 5:
            # Fallback if too few peaks for stable GPD fit
            # We assume an exponential distribution as a safe fallback
            self.c = 0.0
            self.scale = np.mean(peaks) if self.N_q > 0 else np.std(scores)
            return self
            
        # Fit GPD parameters via MLE with location forced to 0
        # scipy genpareto.fit returns: shape (c), location, scale
        try:
            params = genpareto.fit(peaks, floc=0)
            self.c = params[0]
            self.scale = params[2]
        except Exception:
            # Fallback to exponential fit (GPD with c=0)
            self.c = 0.0
            self.scale = np.mean(peaks)
            
        return self

    def calculate_threshold(self, N_total=None):
        """
        Calculates the adaptive extreme value threshold using GPD formulas.
        """
        if self.q is None or self.scale is None:
            raise ValueError("SPOTDetector must be fitted first.")
            
        N = N_total if N_total is not None else self.N
        N_q = self.N_q
        P = self.probability
        
        # Guard against zero division or extreme ratio
        if N_q == 0 or N == 0:
            return self.q
            
        ratio = (N * P) / N_q
        
        # Compute threshold based on shape parameter c
        if abs(self.c) < 1e-9:
            # Exponential limit (c -> 0)
            threshold = self.q - self.scale * np.log(ratio)
        else:
            # General GPD formula
            threshold = self.q + (self.scale / self.c) * (np.power(ratio, -self.c) - 1.0)
            
        return float(threshold)

    def is_anomaly(self, score, N_total=None):
        """
        Determines if a score is anomalous and dynamically updates GPD parameters
        if the score is considered normal (benign streaming update).
        """
        threshold = self.calculate_threshold(N_total)
        if score >= threshold:
            return True, threshold
            
        # Benign streaming update: accumulate non-anomalous score
        self.N += 1
        if score > self.q:
            self.peaks.append(score - self.q)
            self.N_q += 1
            # Re-fit GPD dynamically with the new peak
            try:
                params = genpareto.fit(self.peaks, floc=0)
                self.c = params[0]
                self.scale = params[2]
            except Exception:
                pass
                
        return False, threshold


class DSPOTDetector(SPOTDetector):
    """
    Drift Streaming Peak-Over-Threshold (DSPOT) detector.
    Model excesses after detrending / scaling scores using a sliding local window,
    making it robust to systemic shifts and diurnal drift.
    """
    def __init__(self, probability=1e-4, q_percentile=0.98, window_size=20):
        super().__init__(probability, q_percentile)
        self.window_size = window_size
        self.history = []
        self.local_means = []

    def fit(self, baseline_scores):
        scores = np.asarray(baseline_scores)
        self.history = list(scores[-self.window_size:])
        
        # Calculate residuals (detrended scores) using local moving average
        residuals = []
        for i in range(len(scores)):
            start = max(0, i - self.window_size + 1)
            local_mean = np.mean(scores[start:i+1])
            residuals.append(scores[i] - local_mean)
            
        self.local_means = [np.mean(scores[max(0, i - self.window_size + 1):i+1]) for i in range(len(scores))]
        
        # Fit parent SPOT on the residuals
        super().fit(residuals)
        return self

    def calculate_threshold(self, N_total=None, current_local_mean=None):
        """
        Calculates the local adaptive threshold including the current local mean/drift trend.
        """
        residual_threshold = super().calculate_threshold(N_total)
        
        mean_offset = current_local_mean
        if mean_offset is None:
            mean_offset = np.mean(self.history) if self.history else 0.0
            
        return float(mean_offset + residual_threshold)

    def is_anomaly(self, score, N_total=None):
        """
        Evaluates detrended residuals to detect anomalies.
        Updates sliding history and GPD parameters for streaming normal scores.
        """
        local_mean = np.mean(self.history) if self.history else 0.0
        residual = score - local_mean
        
        # Determine anomaly based on residual threshold
        threshold = self.calculate_threshold(N_total, local_mean)
        
        if score >= threshold:
            # Flagged as anomaly
            return True, threshold
            
        # Benign streaming update: update sliding history and residuals
        self.history.append(score)
        if len(self.history) > self.window_size:
            self.history.pop(0)
            
        self.N += 1
        if residual > self.q:
            self.peaks.append(residual - self.q)
            self.N_q += 1
            try:
                params = genpareto.fit(self.peaks, floc=0)
                self.c = params[0]
                self.scale = params[2]
            except Exception:
                pass
                
        return False, threshold
