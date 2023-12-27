import math as math
import xlwings as xw
from datetime import timedelta
from Node import Node
from typing import Tuple, Optional

class Tree:
    """
    Classe pour construire et manipuler un arbre trinomial pour la tarification d'options.
    """

    def __init__(self, root: Node, seuil: float, market, model):
        """
        Initialise l'arbre trinomial.

        Args:
            root (Node): Le nœud racine de l'arbre.
            seuil (float): Seuil utilisé pour la poda de l'arbre.
            market: Objet contenant les données du marché.
            model: Modèle utilisé pour la tarification.
        """
        self.root = root
        self.seuil = seuil
        self.market = market
        self.model = model

    def df(self) -> float:
        """
        Calcule le facteur d'actualisation.

        Returns:
            float: Le facteur d'actualisation.
        """
        return math.exp(-self.market.r * self.model.delta_t)

    def build_tree(self, output: str = "S", print_tree: bool = False, ws: Optional[xw.main.Sheet] = None):
        """
        Construit l'arbre trinomial.

        Args:
            output (str): Le type de sortie (par exemple, prix du sous-jacent).
            print_tree (bool): Indique si l'arbre doit être imprimé dans Excel.
            ws (Optional[xw.main.Sheet]): La feuille Excel pour l'affichage.
        """
        sum_ptotal_before = True

        # Boucle qui permet de construire l'abre de gauche à droite
        for i in range(0, self.model.nbsteps):
            sum_ptotal_before = self.build_nodes_columns(i + 1, self.have_div(i), sum_ptotal_before, print_tree, ws,
                                                         output)
            self.root = self.root.n_mid
        # Afficher l'arbre dans excel
        if print_tree:
            self.last_columns(i + 2, ws, output)
            print(print_tree)


    def have_div(self, i: int) -> bool:
        """
        Détermine si un dividende est dû à un moment donné.

        Args:
            i (int): L'indice de temps.

        Returns:
            bool: Vrai si un dividende est dû, faux sinon.
        """
        # Vérification si la date d'ex-dividende se trouve entre les bornes
        return (self.model.prdate + timedelta(days=i * self.model.delta_t * 365) < self.market.div_date <=
                self.model.prdate + timedelta(days=(i + 1) * self.model.delta_t * 365))

    def last_columns(self, i: int, ws: xw.main.Sheet, output: str):
        """
        Remplit les dernières colonnes de l'arbre dans Excel.

        Args:
            i (int): L'indice de la colonne.
            ws (xw.main.Sheet): La feuille Excel pour l'affichage.
            output (str): Le type de sortie.
        """

        row0 = self.model.nbsteps + 1
        row = row0
        nodes_in_column = 1

        ws.range((row, i)).value = getattr(self.root, output, "Output don't match with Node parameter")

        nbis = self.root
        # Tant que les noeuds du dessus existe, il construit leur block respectif
        while nbis.up is not None:
            nbis = nbis.up
            nodes_in_column += 1
            row -= 1
            ws.range((row, i)).value = getattr(nbis, output, "Output don't match with Node parameter")

        row = row0
        nbis = self.root
        # Tant que les noeuds du dessous existe, il construit leur block respectif
        while nbis.down is not None:
            nbis = nbis.down
            nodes_in_column += 1
            row += 1
            ws.range((row, i)).value = getattr(nbis, output, "Output don't match with Node parameter")


    def build_nodes_columns(self, i: int, have_div: bool, sum_ptotal_before: bool, print_tree: bool,
                            ws: Optional[xw.main.Sheet], output: str) -> bool:
        """
        Construit une colonne de nœuds dans l'arbre.

        Args:
            i (int): L'indice de la colonne.
            have_div (bool): Indique si un dividende est dû.
            sum_ptotal_before (bool): La somme des probabilités totales avant la construction.
            print_tree (bool): Indique si l'arbre doit être imprimé dans Excel.
            ws (Optional[xw.main.Sheet]): La feuille Excel pour l'affichage.
            output (str): Le type de sortie.

        Returns:
            bool: Vrai si la somme des probabilités est correcte, faux sinon.
        """
        sum_ptotal = 0

        row0 = self.model.nbsteps + 1

        row = row0

        # En partant du noeud du milieu, construction du prochain block
        self.root.n_mid = Node(self.root.forward_mid(have_div), i, self.market, self.model)
        self.root.build_block(self.root.n_mid, self, have_div)
        sum_ptotal += self.root.p_total
        nodes_in_column = 1

        if print_tree:
            ws.range((row, i)).value = getattr(self.root, output, "Output don't match with Node parameter")

        nbis = self.root
        # Tant que les noeuds du dessus existe, il construit leur block respectif
        while nbis.up is not None:
            nbis.up.build_block(nbis.n_up, self, have_div)
            sum_ptotal += nbis.up.p_total
            nbis = nbis.up
            nodes_in_column += 1

            if print_tree:
                row -= 1
                ws.range((row, i)).value = getattr(nbis, output, "Output don't match with Node parameter")

        row = row0
        nbis = self.root
        # Tant que les noeuds du dessous existe, il construit leur block respectif
        while nbis.down is not None:
            nbis.down.build_block(nbis.n_down, self, have_div)
            sum_ptotal += nbis.down.p_total
            nbis = nbis.down
            nodes_in_column += 1

            if print_tree:
                row += 1
                ws.range((row, i)).value = getattr(nbis, output, "Output don't match with Node parameter")

        # check si pour chaque step effectué les probas totales sont bien égale 1
        return sum_ptotal - 1 < 10 ** -10 and sum_ptotal_before

    @staticmethod
    def clear_sheet_arbre(ws: xw.main.Sheet):
        """
        Efface complètement le contenu de la feuille Excel spécifiée pour permettre
        l'affichage d'un nouvel arbre.

        Args:
            ws (xw.main.Sheet): La feuille Excel à effacer.
        """
        # Désactive la mise à jour de l'écran et le calcul automatique pour améliorer les performances
        app = xw.apps.active
        app.screen_updating = False
        app.calculation = 'manual'

        try:
            # Efface tous les contenus de la feuille
            ws.clear_contents()
        finally:
            # Réactive la mise à jour de l'écran et le calcul automatique
            app.screen_updating = True
            app.calculation = 'automatic'