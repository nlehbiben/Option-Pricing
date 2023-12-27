from datetime import datetime, timedelta
import numpy as np
from Convergence import Convergence
from ExcelInterface import ExcelInterface

class GreeksCalculator:
    def __init__(self, convergence: Convergence):
        self.convergence = convergence
        self.convergence.data['print_arbre'] = False  # Désactive l'affichage de l'arbre par défaut
    
    def calculate_delta(self, bump: float = 0.01) -> float:
        self.convergence.print_arbre = False
        original_s0 = self.convergence.data['s0']
        bumped_s0 = original_s0 * (1 + bump)
    
        # Sauvegarder la valeur originale de s0
        original_s0_value = self.convergence.data['s0']
    
        try:
            self.convergence.data['s0'] = bumped_s0
            self.convergence.print_arbre = False  # Assurez-vous que print_arbre est correctement défini ici
            bumped_price = self.convergence.run_trinomial()
        finally:
            # Restaurer la valeur originale de s0 et print_arbre
            self.convergence.data['s0'] = original_s0_value
            self.convergence.print_arbre = True  # Restaurez la valeur précédente de print_arbre
    
        original_price = self.convergence.run_trinomial()
    
        delta = (bumped_price - original_price) / (bump * original_s0)
        return delta

    
    def calculate_gamma(self, bump: float = 0.01) -> float:
        self.convergence.print_arbre = False
        original_s0 = self.convergence.data['s0']
        bumped_s0_up = original_s0 * (1 + bump)
        bumped_s0_down = original_s0 * (1 - bump)
    
        # Sauvegarder la valeur originale de s0 et de print_arbre
        original_s0_value = self.convergence.data['s0']
        original_print_arbre = self.convergence.print_arbre
    
        try:
            self.convergence.data['s0'] = bumped_s0_up
            self.convergence.print_arbre = False
            price_up = self.convergence.run_trinomial()
    
            self.convergence.data['s0'] = bumped_s0_down
            price_down = self.convergence.run_trinomial()
    
            self.convergence.data['s0'] = original_s0
            original_price = self.convergence.run_trinomial()
    
            gamma = (price_up - 2 * original_price + price_down) / ((bump * original_s0) ** 2)
            return gamma
        finally:
            # Restaurer la valeur originale de s0 et de print_arbre
            self.convergence.data['s0'] = original_s0_value
            self.convergence.print_arbre = original_print_arbre

    
    def calculate_vega(self, volatility_increment: float = 0.01) -> float:
        self.convergence.print_arbre = False
        original_vol = self.convergence.data['vol']
    
        # Sauvegarder la valeur originale de vol et de print_arbre
        original_vol_value = self.convergence.data['vol']
        original_print_arbre = self.convergence.print_arbre
    
        try:
            self.convergence.data['vol'] += volatility_increment
            self.convergence.print_arbre = False
            price_up = self.convergence.run_trinomial()
    
            self.convergence.data['vol'] = original_vol
            original_price = self.convergence.run_trinomial()
    
            vega = (price_up - original_price) / volatility_increment
            return vega
        finally:
            # Restaurer la valeur originale de vol et de print_arbre
            self.convergence.data['vol'] = original_vol_value
            self.convergence.print_arbre = original_print_arbre

    
    def calculate_rho(self, interest_increment: float = 0.01) -> float:
        self.convergence.print_arbre = False
        original_rate = self.convergence.data['r']
    
        # Sauvegarder la valeur originale de r et de print_arbre
        original_rate_value = self.convergence.data['r']
        original_print_arbre = self.convergence.print_arbre
    
        try:
            self.convergence.data['r'] += interest_increment
            self.convergence.print_arbre = False
            price_up = self.convergence.run_trinomial()
    
            self.convergence.data['r'] = original_rate
            original_price = self.convergence.run_trinomial()
    
            rho = (price_up - original_price) / interest_increment
            return rho
        finally:
            # Restaurer la valeur originale de r et de print_arbre
            self.convergence.data['r'] = original_rate_value
            self.convergence.print_arbre = original_print_arbre

    
    def calculate_theta(self, time_decrement_days: float = 0.01) -> float:
        self.convergence.print_arbre = False
        original_maturity = self.convergence.data['maturity']
    
        # Sauvegarder la valeur originale de maturity et de print_arbre
        original_maturity_value = self.convergence.data['maturity']
        original_print_arbre = self.convergence.print_arbre
    
        try:
            self.convergence.print_arbre = False
            time_decrement = timedelta(days=time_decrement_days)
    
            if isinstance(original_maturity, str):
                original_maturity = datetime.strptime(original_maturity, '%Y-%m-%d')
    
            new_maturity = original_maturity - time_decrement
            self.convergence.data['maturity'] = new_maturity.strftime('%Y-%m-%d')
            price_down = self.convergence.run_trinomial()
    
            self.convergence.data['maturity'] = original_maturity_value
            original_price = self.convergence.run_trinomial()
    
            theta = (price_down - original_price) / time_decrement_days
            return theta
        finally:
            # Restaurer la valeur originale de maturity et de print_arbre
            self.convergence.data['maturity'] = original_maturity_value
            self.convergence.print_arbre = original_print_arbre

    
    def Graph_delta(self, excel_interface: ExcelInterface) -> None:
        """
        Calcule le Delta pour une plage de valeurs du sous-jacent autour du strike et exporte dans Excel.

        Args:
            excel_interface (ExcelInterface): L'interface pour interagir avec Excel.
        """
        # Obtenir le prix d'exercice de l'option
        strike = self.convergence.data['strike']

        # Gamme de prix du sous-jacent de 0 à 2 fois le strike, par pas de 1
        s0_range = np.arange(1, strike * 2.00 + 1, 1)

        # Liste pour stocker les résultats du Delta
        delta_results = [('Sous-jacent', 'Delta')]  # En-tête pour les colonnes

        # Calcul de Delta pour chaque prix du sous-jacent
        for s0 in s0_range:
            self.convergence.data['s0'] = s0
            delta = round(self.calculate_delta(), 6)
            delta_results.append((s0, delta))

        # Exportation des résultats dans Excel
        excel_interface.wb.sheets['Greeks'].range('A1').value = delta_results

    def Graph_gamma(self, excel_interface: ExcelInterface) -> None:
        """
        Calcule le Gamma pour une plage de valeurs du sous-jacent autour du strike et exporte dans Excel.
        """
        strike = self.convergence.data['strike']
        s0_range = np.arange(0, strike * 2.00 + 1, 1)

        gamma_results = [('Sous-jacent', 'Gamma')]

        for s0 in s0_range:
            self.convergence.data['s0'] = s0
            gamma = round(self.calculate_gamma(), 6)
            gamma_results.append((s0, gamma))

        excel_interface.wb.sheets['Greeks'].range('B1').value = gamma_results

    def Graph_vega(self, excel_interface: ExcelInterface) -> None:
        """
        Calcule le Vega pour une plage de valeurs du sous-jacent autour du strike et exporte dans Excel.
        """
        strike = self.convergence.data['strike']
        s0_range = np.arange(0, strike * 2.00 + 1, 1)

        vega_results = [('Sous-jacent', 'Vega')]

        for s0 in s0_range:
            self.convergence.data['s0'] = s0
            vega = round(self.calculate_vega(), 6)
            vega_results.append((s0, vega))

        excel_interface.wb.sheets['Greeks'].range('C1').value = vega_results

    def Graph_theta(self, excel_interface: ExcelInterface) -> None:
        """
        Calcule le Theta pour une plage de valeurs du sous-jacent autour du strike et exporte dans Excel.
        """
        strike = self.convergence.data['strike']
        s0_range = np.arange(0, strike * 2.00 + 1, 1)

        theta_results = [('Sous-jacent', 'Theta')]

        for s0 in s0_range:
            self.convergence.data['s0'] = s0
            theta = round(self.calculate_theta(), 6)
            theta_results.append((s0, theta))

        excel_interface.wb.sheets['Greeks'].range('E1').value = theta_results

    def Graph_rho(self, excel_interface: ExcelInterface) -> None:
        """
        Calcule le Rho pour une plage de valeurs du sous-jacent autour du strike et exporte dans Excel.
        """
        strike = self.convergence.data['strike']
        s0_range = np.arange(0, strike * 2.00 + 1, 1)

        rho_results = [('Sous-jacent', 'Rho')]

        for s0 in s0_range:
            self.convergence.data['s0'] = s0
            rho = round(self.calculate_rho(), 6)
            rho_results.append((s0, rho))

        excel_interface.wb.sheets['Greeks'].range('G1').value = rho_results