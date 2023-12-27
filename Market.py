class Market:
    """
    La classe Market définit l'environnement de marché pour une option.

    Attributs:
        r (float): Taux d'intérêt sans risque.
        vol (float): Volatilité du sous-jacent.
        s0 (float): Prix initial du sous-jacent.
        div (float): Dividende.
        div_date (date): Date d'ex-dividende.
    """

    def __init__(self, r, vol, s0, div, div_date):
        """
        Initialise une nouvelle instance de la classe Market.

        Args:
            r (float): Taux d'intérêt sans risque.
            vol (float): Volatilité du sous-jacent.
            s0 (float): Prix initial du sous-jacent.
            div (float): Dividende.
            div_date (date): Date d'ex-dividende.
        """
        self.r = r
        self.vol = vol
        self.s0 = s0
        self.div = div
        self.div_date = div_date
