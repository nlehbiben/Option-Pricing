from datetime import datetime

class Option:
    """
    La classe Option définit les caractéristiques d'une option.
    """

    def __init__(self, option_type, type, strike, maturity):
        """
        Initialise une nouvelle instance de la classe Option.
        """
        self.op_type = option_type
        self.type = type
        self.strike = strike
        self.maturity = maturity

    def payoff(self, spot) -> float:
        """
        Calcule le paiement de l'option en fonction du prix du sous-jacent à maturité.
        """
        if self.op_type == "Call":
            return max(spot - self.strike, 0)
        else:
            return max(self.strike - spot, 0)
