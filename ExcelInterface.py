import xlwings as xw
from typing import Dict, Any, List, Tuple, Optional


class ExcelInterface:
    """
    Interface pour interagir avec un classeur Excel en utilisant xlwings.
    Gère la lecture des données et l'écriture des résultats dans des feuilles de calcul spécifiques.
    """

    def __init__(self, workbook_path: str):
        """
        Initialise l'ExcelInterface avec le chemin vers le classeur.

        Args:
            workbook_path (str): Chemin d'accès au fichier Excel.
        """
        self.wb = xw.Book(workbook_path)
        self.sht = self.wb.sheets['Pricing']
        self.sht_conv_nbsteps = self.wb.sheets['Convergence_NbSteps']
        self.sht_conv_strike = self.wb.sheets['Convergence_Strike']
        self.sht_arbre = self.wb.sheets['Arbre']

    def read_data(self) -> Dict[str, Any]:
        """
        Lit les données de configuration à partir de la feuille Excel 'Pricing'.

        Returns:
            Dict[str, Any]: Un dictionnaire contenant les paramètres lus.
        """
        return {
            'r': self.sht.range('Rate').value,
            'vol': self.sht.range('Vol').value,
            's0': self.sht.range('StartPrice').value,
            'div': self.sht.range('Div').value,
            'div_date': self.sht.range('Div_Date').value,
            'option_type': self.sht.range('Type').value,
            'type': self.sht.range('Type_exercice').value,
            'strike': self.sht.range('Strike').value,
            'maturity': self.sht.range('Maturity').value,
            'pricing_date': self.sht.range('AsOfDate').value,
            'nbsteps': int(self.sht.range('Nb_Steps').value),
            'max_steps': int(self.sht.range('max_steps').value),
            'print_arbre': bool(self.sht.range('Print_tree').value),
            'is_pruned': self.sht.range('is_pruning').value.strip(),
            'pruned_level': self.sht.range('level_pruning').value
        }

    def write_trinomial_results(self, trinomial_result: Optional[float] = None,
                                trinomial_greeks: Optional[Dict[str, float]] = None):
        """
        Écrit les résultats du modèle trinomial dans la feuille Excel.
    
        Args:
            trinomial_result (Optional[float]): Prix calculé par le modèle trinomial.
            trinomial_greeks (Optional[Dict[str, float]]): Grecques du modèle trinomial.
        """
        euro_symbol = '€'
        decimal_places = 6  # Nombre de décimales après la virgule
    
        # Définition des cellules de résultat
        trinomial_price_cell = 'H3'
        greeks_cells = {
            'Delta': 'H8',
            'Gamma': 'H10',
            'Vega': 'H12',
            'Rho': 'H14',
            'Theta': 'H16'
        }
    
        if trinomial_result is not None:
            formatted_trinomial_result = f"{round(trinomial_result, decimal_places):.{decimal_places}f} {euro_symbol}"
            self.sht.range(trinomial_price_cell).value = formatted_trinomial_result
    
        if trinomial_greeks:
            for greek, cell in greeks_cells.items():
                if greek in trinomial_greeks:
                    self.sht.range(cell).value = f"{trinomial_greeks[greek]:.{decimal_places}f}"
                    
    def write_black_scholes_results(self, bs_result: Optional[Dict[str, float]] = None):
        """
        Écrit les résultats du modèle Black & Scholes dans la feuille Excel.
    
        Args:
            bs_result (Optional[Dict[str, float]]): Résultats du modèle Black & Scholes.
        """
        euro_symbol = '€'
        decimal_places = 6  # Nombre de décimales après la virgule
    
        # Définition des cellules de résultat
        bs_price_cell = 'J3'
        bs_greeks_cells = {
            'Delta': 'J8',
            'Gamma': 'J10',
            'Vega': 'J12',
            'Rho': 'J14',
            'Theta': 'J16'
        }
    
        if bs_result and 'Greeks' in bs_result and bs_result['Greeks'] is not None:
            self.sht.range(bs_price_cell).value = f"{bs_result['Price']:.{decimal_places}f} {euro_symbol}"
            for greek, cell in bs_greeks_cells.items():
                if greek in bs_result['Greeks']:
                    self.sht.range(cell).value = f"{bs_result['Greeks'][greek]:.{decimal_places}f}"



    def write_nbsteps_convergence_results(self, convergence_results: List[Tuple[int, float, float, float]]):
        """
        Écrit les résultats de convergence en fonction du nombre d'étapes dans la feuille Excel dédiée.

        Args:
            convergence_results (List[Tuple[int, float, float, float]]): Résultats de convergence.
        """
        # Définition des cellules pour les titres et les données
        header_cell = 'A1'
        data_start_cell = 'A2'

        # Écriture des titres des colonnes
        self.sht_conv_nbsteps.range(header_cell).value = [
            'NbSteps', 'TrinomialPrice', 'BlackScholesPrice', '(Trinomial - Black_Scholes) x NbSteps'
        ]

        # Écriture des données
        self.sht_conv_nbsteps.range(data_start_cell).value = convergence_results

    def write_strike_convergence_results(self, convergence_results: List[Tuple[int, float, float, float, float, float]]):
        """
        Écrit les résultats de convergence en fonction du prix d'exercice (strike) dans la feuille Excel dédiée.

        Args:
            convergence_results (List[Tuple[int, float, float, float, float, float]]): Résultats de convergence.
        """
        # Définition des cellules pour les titres et les données
        header_cell = 'A1'
        data_start_cell = 'A2'

        # Écriture des titres des colonnes
        self.sht_conv_strike.range(header_cell).value = [
            'Strike', 'TrinomialPrice', 'BlackScholesPrice', 'Trinomial - Black_Scholes', 'Trinomial Slope', 'BS Slope'
        ]

        # Écriture des données
        self.sht_conv_strike.range(data_start_cell).value = convergence_results
