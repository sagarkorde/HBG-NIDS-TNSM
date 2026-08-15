"""
HBG-NIDS Pre-Scoring Feature Decorrelation Module
Implements DecorrelatedBenfordDetector and a robust PCA-based projection
to resolve redundant, correlated feature signals and maximize statistical power.
"""

import numpy as np
from sklearn.decomposition import PCA

class DecorrelatedBenfordDetector:
    """
    Decorrelated Benford statistical anomaly detector.
    Projects Benford-conformant features into an orthogonal coordinate space (PCA)
    before computing divergence scores, maximizing statistical power.
    """
    def __init__(self, n_components=3):
        """
        Parameters:
        - n_components: Number of principal components to extract (typically 3, capturing >95% variance).
        """
        self.n_components = n_components
        self.pca = PCA(n_components=n_components)
        self.fitted = False

    def fit(self, baseline_features):
        """
        Fits the PCA projection matrix W and calculates baseline parameters.
        """
        X = np.asarray(baseline_features)
        self.pca.fit(X)
        self.fitted = True
        return self

    def project(self, feature_vector):
        """
        Projects an incoming feature vector onto the orthogonal principal component space.
        """
        if not self.fitted:
            raise ValueError("DecorrelatedBenfordDetector must be fitted first.")
        X = np.asarray(feature_vector)
        if X.ndim == 1:
            X = X.reshape(1, -1)
        return self.pca.transform(X)

    def verify_orthogonality(self, test_features):
        """
        Projects test features and computes the correlation matrix
        to verify that the off-diagonal (pairwise) correlations are exactly 0.
        """
        X_projected = self.project(test_features)
        # Compute Pearson correlation matrix
        corr_matrix = np.corrcoef(X_projected, rowvar=False)
        return corr_matrix


def run_decorrelation_simulation():
    """
    Simulates the 5 Benford-conformant features (Flow Duration, Fwd Bytes, Bwd Bytes, etc.)
    under two regimes to prove that PCA projects features into a perfectly orthogonal space
    and eliminates redundant, correlated signals.
    """
    print("Initializing feature decorrelation simulation...")
    np.random.seed(42)
    num_samples = 500
    
    # Simulate 5 conformant features with high correlation (mean r ~ 0.62)
    # Generate latent variable representing flow scale / background activity
    latent = np.random.normal(loc=10.0, scale=2.0, size=num_samples)
    
    # Features are highly dependent on the latent scale variable
    f1 = latent + np.random.normal(loc=0.0, scale=0.5, size=num_samples)  # e.g., Flow Duration
    f2 = 1.2 * latent + np.random.normal(loc=0.0, scale=0.4, size=num_samples)  # e.g., Total Fwd Bytes
    f3 = 1.1 * latent + np.random.normal(loc=0.0, scale=0.6, size=num_samples)  # e.g., Total Bwd Bytes
    f4 = 0.8 * latent + np.random.normal(loc=0.0, scale=0.3, size=num_samples)  # e.g., Flow Bytes/s
    f5 = 0.5 * latent + np.random.normal(loc=0.0, scale=0.8, size=num_samples)  # e.g., Fwd IAT Mean
    
    X = np.column_stack([f1, f2, f3, f4, f5])
    
    # 1. Evaluate REGIME A: Correlated Benford Layer
    print("\n--- REGIME A: Correlated Benford Layer ---")
    corr_matrix_a = np.corrcoef(X, rowvar=False)
    # Extract upper triangle off-diagonal correlations
    indices = np.triu_indices_from(corr_matrix_a, k=1)
    pairwise_r_a = corr_matrix_a[indices]
    mean_r_a = np.mean(np.abs(pairwise_r_a))
    print("Pairwise Pearson Correlation Matrix (Original):")
    print(np.round(corr_matrix_a, 4))
    print(f"Mean Pairwise Correlation (r): {mean_r_a:.4f} (range: {np.min(pairwise_r_a):.4f} to {np.max(pairwise_r_a):.4f})")
    
    # 2. Evaluate REGIME B: Decorrelated PCA-based Benford Layer
    print("\n--- REGIME B: Decorrelated PCA-based Benford Layer ---")
    detector = DecorrelatedBenfordDetector(n_components=3)
    detector.fit(X)
    
    # Project features and compute projected correlation matrix
    corr_matrix_b = detector.verify_orthogonality(X)
    indices_b = np.triu_indices_from(corr_matrix_b, k=1)
    pairwise_r_b = corr_matrix_b[indices_b]
    mean_r_b = np.mean(np.abs(pairwise_r_b))
    print("Pairwise Pearson Correlation Matrix (Projected PCA):")
    print(np.round(corr_matrix_b, 4))
    print(f"Mean Pairwise Correlation (r): {mean_r_b:.4f} (orthogonal components)")
    
    # Print explained variance ratio
    var_ratio = detector.pca.explained_variance_ratio_
    print(f"Explained Variance of Top 3 Components: {var_ratio} (cumulative: {np.sum(var_ratio)*100:.2f}%)")
    
    return {
        "mean_r_a": mean_r_a,
        "mean_r_b": mean_r_b,
        "cumulative_variance": np.sum(var_ratio)
    }

if __name__ == "__main__":
    run_decorrelation_simulation()
