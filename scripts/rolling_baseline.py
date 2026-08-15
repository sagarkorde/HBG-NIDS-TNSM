"""
HBG-NIDS Weekday-Adaptive Rolling Baseline Module
Implements WeekdayBaselineModel and a robust simulation engine to
quantify the reduction of Thursday false positives under temporal drift.
"""

import numpy as np

class WeekdayBaselineModel:
    """
    Weekday-Adaptive Rolling Baseline tracker.
    Maintains separate rolling historical distributions for each day of the week
    to scale incoming flow records and mitigate day-of-week and diurnal traffic drift.
    """
    def __init__(self, num_features, history_limit=5):
        """
        Parameters:
        - num_features: Dimension of the flow metadata feature vector.
        - history_limit: Number of historical days of the same weekday to retain.
        """
        self.num_features = num_features
        self.history_limit = history_limit
        # Partition history by day index (0=Monday, ..., 6=Sunday)
        self.history = {day: [] for day in range(7)}
        self.means = {day: np.zeros(num_features) for day in range(7)}
        self.stds = {day: np.ones(num_features) for day in range(7)}

    def seed_baseline(self, day, daily_scores):
        """
        Seeds baseline history for a specific day with known benign data.
        """
        scores = np.asarray(daily_scores)
        if scores.ndim == 1:
            scores = scores.reshape(-1, 1)
        self.history[day].append(scores)
        self._update_stats(day)

    def _update_stats(self, day):
        """
        Recomputes means and standard deviations for a specific day.
        """
        all_samples = np.vstack(self.history[day])
        self.means[day] = np.mean(all_samples, axis=0)
        self.stds[day] = np.std(all_samples, axis=0)
        # Avoid division by zero
        self.stds[day][self.stds[day] < 1e-8] = 1e-8

    def scale(self, day, feature_vector):
        """
        Scales an incoming feature vector using weekday-specific historical statistics.
        """
        x = np.asarray(feature_vector)
        z = (x - self.means[day]) / self.stds[day]
        return z

    def update(self, day, feature_vector):
        """
        Appends normal/benign feature vector to sliding history.
        """
        x = np.asarray(feature_vector)
        if x.ndim == 1:
            x = x.reshape(1, -1)
            
        self.history[day].append(x)
        if len(self.history[day]) > self.history_limit:
            self.history[day].pop(0)
            
        self._update_stats(day)


def run_drift_simulation():
    """
    Simulates Monday-to-Friday traffic profiles under two regimes:
    Regime A: Monday-only baseline.
    Regime B: Weekday-adaptive rolling baseline.
    
    Verifies that Regime B reduces Thursday false alarm rates (FPR) while preserving
    detection capability for Friday's true volumetric attacks.
    """
    print("Initializing baseline drift simulation...")
    np.random.seed(42)
    
    # 1. Generate diurnal baseline template (normal office hour peak)
    # 24 hours * 4 windows/hour = 96 windows/day. Let's simplify to 34 windows/day
    windows_per_day = 34
    num_features = 3  # e.g., Flow Duration, Fwd Bytes, Bwd Bytes
    
    # Monday benign baseline
    monday_base = np.random.lognormal(mean=1.0, sigma=0.2, size=(windows_per_day, num_features))
    # Tuesday benign baseline (similar to Monday)
    tuesday_base = np.random.lognormal(mean=1.05, sigma=0.2, size=(windows_per_day, num_features))
    # Wednesday benign baseline
    wednesday_base = np.random.lognormal(mean=1.02, sigma=0.2, size=(windows_per_day, num_features))
    
    # Thursday normal drift (company-wide update / natural weekly peak: +80% volume increase)
    thursday_drift = np.random.lognormal(mean=1.8, sigma=0.2, size=(windows_per_day, num_features))
    
    # Friday normal baseline
    friday_base = np.random.lognormal(mean=1.0, sigma=0.2, size=(windows_per_day, num_features))
    # Friday DDoS True Positives (massive out-of-distribution spike: +600% increase)
    friday_ddos = np.random.lognormal(mean=7.0, sigma=0.5, size=(10, num_features))
    
    # Assemble complete Friday scores (24 benign windows, 10 DDoS attack windows)
    friday_scores = np.vstack([friday_base[:24], friday_ddos])
    
    # Initialize baseline models
    static_mean = np.mean(monday_base, axis=0)
    static_std = np.std(monday_base, axis=0)
    
    adaptive_model = WeekdayBaselineModel(num_features=num_features, history_limit=3)
    # Seed historical weekday profiles
    adaptive_model.seed_baseline(0, monday_base)
    adaptive_model.seed_baseline(1, tuesday_base)
    adaptive_model.seed_baseline(2, wednesday_base)
    # Suppose we have historical Thursdays that match the Thursday drift profile
    historical_thursdays = np.random.lognormal(mean=1.78, sigma=0.2, size=(windows_per_day, num_features))
    adaptive_model.seed_baseline(3, historical_thursdays)
    # Seed Friday history
    historical_fridays = np.random.lognormal(mean=1.01, sigma=0.2, size=(windows_per_day, num_features))
    adaptive_model.seed_baseline(4, historical_fridays)
    
    # 2. Evaluate REGIME A (Static Monday-Only Baseline)
    print("\n--- REGIME A: Static Monday-Only Baseline ---")
    threshold = 2.5  # standard outlier boundary
    
    # Thursday Evaluation under Regime A
    thu_scaled_a = (thursday_drift - static_mean) / static_std
    thu_alerts_a = np.any(thu_scaled_a > threshold, axis=1)
    thu_fpr_a = np.sum(thu_alerts_a) / windows_per_day
    print(f"Thursday Benign FP Rate: {thu_fpr_a*100:.1f}% ({np.sum(thu_alerts_a)}/{windows_per_day} alerts)")
    
    # Friday DDoS Evaluation under Regime A
    fri_scaled_a = (friday_scores - static_mean) / static_std
    ddos_alerts_a = np.any(fri_scaled_a[24:] > threshold, axis=1)
    ddos_recall_a = np.sum(ddos_alerts_a) / 10
    print(f"Friday DDoS Recall: {ddos_recall_a*100:.1f}%")
    
    # 3. Evaluate REGIME B (Weekday-Adaptive Rolling Baseline)
    print("\n--- REGIME B: Weekday-Adaptive Rolling Baseline ---")
    
    # Thursday Evaluation under Regime B
    thu_scaled_b = np.array([adaptive_model.scale(3, x) for x in thursday_drift])
    thu_alerts_b = np.any(thu_scaled_b > threshold, axis=1)
    thu_fpr_b = np.sum(thu_alerts_b) / windows_per_day
    print(f"Thursday Benign FP Rate: {thu_fpr_b*100:.1f}% ({np.sum(thu_alerts_b)}/{windows_per_day} alerts)")
    
    # Friday DDoS Evaluation under Regime B
    fri_scaled_b = np.array([adaptive_model.scale(4, x) for x in friday_scores])
    ddos_alerts_b = np.any(fri_scaled_b[24:] > threshold, axis=1)
    ddos_recall_b = np.sum(ddos_alerts_b) / 10
    print(f"Friday DDoS Recall: {ddos_recall_b*100:.1f}%")
    
    return {
        "thu_fpr_a": thu_fpr_a,
        "thu_fpr_b": thu_fpr_b,
        "ddos_recall_a": ddos_recall_a,
        "ddos_recall_b": ddos_recall_b
    }

if __name__ == "__main__":
    run_drift_simulation()
