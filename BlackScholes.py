import math as m
import scipy.stats as stats
from typing import Dict
from datetime import datetime



class BlackScholes:
    def __init__(self, market, option, model) -> None:
        """
        Initialisation de la classe BlackScholes.

        Args:
            market: Instance de la classe Market contenant les informations du marché.
            option: Instance de la classe Option contenant les détails de l'option.
            model: Instance de la classe Model utilisée pour le pricing.
        """
        self.market = market
        self.option = option
        self.model = model

    def d1(self) -> float:
        """
        Calcule le paramètre d1 utilisé dans la formule Black-Scholes.

        Returns:
            float: Valeur calculée de d1.
        """
        t = (self.option.maturity - self.model.prdate).days / 365
        numerator = m.log(self.market.s0 / self.option.strike) + (self.market.r + 0.5 * self.market.vol ** 2) * t
        denominator = self.market.vol * m.sqrt(t)
        return numerator / denominator

    def d2(self) -> float:
        """
        Calcule le paramètre d2 utilisé dans la formule Black-Scholes.

        Returns:
            float: Valeur calculée de d2.
        """
        t = (self.option.maturity - self.model.prdate).days / 365
        return self.d1() - self.market.vol * m.sqrt(t)

    def price(self) -> float:
        """
        Calcule le prix de l'option en utilisant la formule Black-Scholes.

        Returns:
            float: Prix calculé de l'option.
        """
        t = (self.option.maturity - self.model.prdate).days / 365
        if self.option.op_type == "Call":
            return (self.market.s0 * stats.norm.cdf(self.d1()) -
                    self.option.strike * m.exp(-self.market.r * t) * stats.norm.cdf(self.d2()))
        else:
            return (self.option.strike * m.exp(-self.market.r * t) * stats.norm.cdf(-self.d2()) -
                    self.market.s0 * stats.norm.cdf(-self.d1()))

    def calculate_greeks(self) -> Dict[str, float]:
        """
        Calcule les Grecques de l'option.

        Returns:
            Dict[str, float]: Dictionnaire contenant les Grecques calculées.
        """
        delta = self.delta()
        gamma = self.gamma()
        vega = self.vega()
        theta = self.theta()
        rho = self.rho()
        return {
            "Delta": delta,
            "Gamma": gamma,
            "Vega": vega,
            "Theta": theta,
            "Rho": rho
        }

    def delta(self) -> float:
        """
        Calcule la Grecque Delta de l'option.

        Returns:
            float: Valeur calculée de Delta.
        """
        if self.option.op_type == "Call":
            return stats.norm.cdf(self.d1())
        else:
            return -stats.norm.cdf(-self.d1())

    def gamma(self) -> float:
        """
        Calcule la Grecque Gamma de l'option.

        Returns:
            float: Valeur calculée de Gamma.
        """
        t = (self.option.maturity - self.model.prdate).days / 365
        return stats.norm.pdf(self.d1()) / (self.market.s0 * self.market.vol * m.sqrt(t))

    def vega(self) -> float:
        """
        Calcule la Grecque Vega de l'option.

        Returns:
            float: Valeur calculée de Vega.
        """
        t = (self.option.maturity - self.model.prdate).days / 365
        return self.market.s0 * stats.norm.pdf(self.d1()) * m.sqrt(t)

    def theta(self) -> float:
        """
        Calcule la Grecque Theta de l'option.

        Returns:
            float: Valeur calculée de Theta.
        """
        t = (self.option.maturity - self.model.prdate).days / 365
        if self.option.op_type == "Call":
            return -(self.market.s0 * stats.norm.pdf(self.d1()) * self.market.vol) / (2 * m.sqrt(t)) - \
                self.market.r * self.option.strike * m.exp(-self.market.r * t) * stats.norm.cdf(self.d2())
        else:
            return -(self.market.s0 * stats.norm.pdf(self.d1()) * self.market.vol) / (2 * m.sqrt(t)) + \
                self.market.r * self.option.strike * m.exp(-self.market.r * t) * stats.norm.cdf(-self.d2())

    def rho(self) -> float:
        """
        Calcule la Grecque Rho de l'option.

        Returns:
            float: Valeur calculée de Rho.
        """
        t = (self.option.maturity - self.model.prdate).days / 365
        if self.option.op_type == "Call":
            return self.option.strike * t * m.exp(-self.market.r * t) * stats.norm.cdf(self.d2())
        else:
            return -self.option.strike * t * m.exp(-self.market.r * t) * stats.norm.cdf(-self.d2())
