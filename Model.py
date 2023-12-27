import math as m
from Option import Option
from Market import Market
from datetime import datetime


class Model:
    """
    La classe Model configure les spécificités du modèle utilisé pour le pricing.

    Attributs:
        market (Market): Instance de la classe Market définissant l'environnement de marché.
        prdate (date): Date de pricing.
        nbsteps (int): Nombre d'étapes dans le modèle.
        delta_t (float): Intervalle de temps entre les étapes.
        alpha (float): Paramètre alpha utilisé pour ajuster les mouvements de prix.
    """

    def __init__(self, pricing_date, nbsteps, option, market):
        """
        Initialise une nouvelle instance de la classe Model.

        Args:
            pricing_date (date): Date de pricing.
            nbsteps (int): Nombre d'étapes dans le modèle.
            option (Option): Instance de la classe Option.
            market (Market): Instance de la classe Market.
        """

        self.prdate = pricing_date if isinstance(pricing_date, datetime) else datetime.strptime(pricing_date,
        '%Y-%m-%d')
        self.maturity = option.maturity if isinstance(option.maturity, datetime) else datetime.strptime(option.maturity,
        '%Y-%m-%d')

        self.delta_t = (self.maturity - self.prdate).days / 365 / nbsteps
        if self.delta_t <= 0:
            raise ValueError(f"Delta_t doit être positif. Calculé: {self.delta_t}")

        self.market = market
        self.nbsteps = nbsteps
        self.delta_t = (self.maturity - self.prdate).days / 365 / nbsteps
        self.alpha = self.calc_alpha()

    def calc_alpha(self) -> float:
        """
        Calcule le paramètre alpha utilisé pour ajuster les mouvements de prix.

        Returns:
            float: Valeur du paramètre alpha.
        """
        # Assurer que self.delta_t est positif
        if self.delta_t <= 0:
            raise ValueError(f"Delta_t est non positif.")
        return m.exp(self.market.vol * m.sqrt(3 * self.delta_t))