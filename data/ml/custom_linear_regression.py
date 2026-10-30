import numpy as np

class CustomLinearRegression:
    def __init__(self, learning_rate=0.01, epochs=1000):
        self.learning_rate = learning_rate
        self.epochs = epochs

        self.weights = None
        self.bias = None

        # For normalization
        self.mean = None
        self.std = None

        self.loss_history = []

    def _normalize(self, X):
        return (X - self.mean) / self.std

    def fit(self, X, y):
        X = np.array(X, dtype=np.float64)
        y = np.array(y, dtype=np.float64)

        n_samples, n_features = X.shape

        # Compute mean & std for normalization
        self.mean = np.mean(X, axis=0)
        self.std = np.std(X, axis=0)

        # Avoid division by zero
        self.std[self.std == 0] = 1

        # Normalize features
        X = self._normalize(X)

        # Initialize parameters
        self.weights = np.zeros(n_features)
        self.bias = 0.0

        # Gradient Descent
        for _ in range(self.epochs):
            y_pred = np.dot(X, self.weights) + self.bias

            dw = (1 / n_samples) * np.dot(X.T, (y_pred - y))
            db = (1 / n_samples) * np.sum(y_pred - y)

            self.weights -= self.learning_rate * dw
            self.bias -= self.learning_rate * db

            loss = np.mean((y_pred - y) ** 2)
            self.loss_history.append(loss)

    def predict(self, X):
        X = np.array(X, dtype=np.float64)
        X = self._normalize(X)
        return np.dot(X, self.weights) + self.bias
