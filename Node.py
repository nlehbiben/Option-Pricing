import math as m

class Node:
    """
    Classe représentant un nœud dans un arbre trinomial pour la modélisation d'options financières.
    """

    def __init__(self, price: float, i: int, market, model, p_total: float = 0, seuil_enabled: bool = True):
        """
        Initialise un nœud de l'arbre trinomial.

        Args:
            price (float): Prix du sous-jacent à ce nœud.
            i (int): Index du nœud (niveau dans l'arbre).
            market: Objet du marché contenant les données du marché.
            model: Modèle utilisé pour la tarification.
            p_total (float): Probabilité totale de se retrouver à ce nœud depuis la racine de l'arbre.
            seuil_enabled (bool): Indique si le seuil est activé pour la poda de l'arbre.
        """
        self.market = market
        self.model = model
        self.S = price  # Prix du sous-jacent à ce nœud
        self.i = i  # Index du nœud (niveau dans l'arbre)
        self.var = self.variance()  # Variance pour ce nœud
        self.p_total = p_total  # Proba totale de se retrouver à ce nœud depuis la racine de l'arbre
        self.pdown, self.pmid, self.pup = 0, 0, 0  # Initialisation des probabilités de transition
        self.opt_value = None  # Payoff de l'option à ce nœud
        self.n_mid, self.n_up, self.n_down, self.down, self.up = None, None, None, None, None
        self.seuil_enabled = seuil_enabled

    def variance(self) -> float:
        """
        Calcule la variance à ce nœud.

        Returns:
            float: La variance calculée pour ce nœud.
        """
        return ((self.S ** 2) * m.exp(2 * self.market.r * self.model.delta_t) *
                (m.exp(self.market.vol ** 2 * self.model.delta_t) - 1))
    def forward_mid(self, have_div: bool) -> float:
        """
        Calcule le prix forward médian du nœud.

        Args:
            have_div (bool): Indique si un dividende doit être pris en compte.

        Returns:
            float: Le prix forward médian.
        """
        if have_div:
            return self.S * m.exp(self.market.r * self.model.delta_t) - self.market.div
        else:
            return self.S * m.exp(self.market.r * self.model.delta_t)

    # Calcule le prix suivant pour un mouvement à la hausse
    def forward_up(self, have_div) -> float:
        return self.forward_mid(have_div) * self.model.alpha

    # Calcule le prix suivant pour un mouvement à la baisse
    def forward_down(self, have_div) -> float:
        return self.forward_mid(have_div) / self.model.alpha

    def is_close(self, fwn: float) -> bool:
        """
        Détermine si le prix forward est proche du prix spot.

        Args:
            fwn (float): Le prix forward à comparer.

        Returns:
            bool: True si fwn est proche du spot, sinon False.
        """
        return self.S * (1 + (1 / self.model.alpha)) / 2 < fwn < self.S * (1 + self.model.alpha) / 2

    def move_up(self) -> 'Node':
        """
        Déplace le nœud vers le haut (augmentation de prix).

        Returns:
            Node: Le nouveau nœud créé ou existant après le déplacement vers le haut.
        """
        if self.up is None:
            self.up = Node(self.S * self.model.alpha, self.i + 1, self.market, self.model)
        return self.up

    def move_down(self) -> 'Node':
        """
        Déplace le nœud vers le bas (diminution de prix).

        Returns:
            Node: Le nouveau nœud créé ou existant après le déplacement vers le bas.
        """
        if self.down is None:
            self.down = Node(self.S / self.model.alpha, self.i + 1, self.market, self.model)
        return self.down

    def get_mid(self, n: 'Node', have_div) -> 'Node':
        """
        Obtient le nœud médian en fonction du prix forward médian.

        Args:
            n (Node): Le prochain nœud à évaluer.

        Returns:
            Node: Le nœud médian trouvé ou créé.
        """
        fwd = self.forward_mid(have_div)

        if n.is_close(fwd):
            return n
        elif fwd > self.S:
            while not n.is_close(fwd):
                n = n.move_up()
        else:
            while not n.is_close(fwd):
                n = n.move_down()
        return n

    def price(self, option, tree) -> float:
        """
        Calcule le prix de l'option à ce nœud.

        Args:
            option: L'option à évaluer.
            tree: L'arbre trinomial utilisé pour le calcul.

        Returns:
            float: Le prix calculé de l'option.
        """
        if self.opt_value is not None:
            return self.opt_value
        elif self.n_mid is None:
            self.opt_value = option.payoff(self.S)
        else:
            if self.p_total > tree.seuil:
                self.opt_value = (self.pup * self.n_up.price(option, tree) +
                                  self.pmid * self.n_mid.price(option, tree) +
                                  self.pdown * self.n_down.price(option, tree)) * tree.df()
            else:
                self.opt_value = self.pmid * self.n_mid.price(option, tree) * tree.df()

            if option.type == "American":
                immediate_exercise_value = option.payoff(self.S)
                self.opt_value = max(self.opt_value, immediate_exercise_value)

        return self.opt_value

    def proba_transition(self, tree, have_div: bool):
        """
        Calcule les probabilités de transition pour le nœud courant.

        Args:
            tree: L'arbre trinomial utilisé pour le calcul.
            have_div (bool): Indique si un dividende doit être pris en compte.
        """
        if self.p_total > tree.seuil:

            self.pdown = (1/self.n_mid.S ** 2 * (self.var + self.forward_mid(have_div) ** 2) - 1 -
                          (self.model.alpha + 1) * (
                    self.forward_mid(have_div)/self.n_mid.S - 1)) / ((1 - self.model.alpha) *
                                                                     (self.model.alpha ** (-2) - 1))
            self.pup = ((self.forward_mid(have_div)/self.n_mid.S - 1 - self.pdown * (1/self.model.alpha - 1))/
                        (self.model.alpha - 1))
            self.pmid = 1 - self.pdown - self.pup
            # Vérifier que les probabilités ne sont pas négatives
            if self.pdown < 0 or self.pup < 0 or self.pmid < 0:
                raise ValueError("Les probabilités de transition ne peuvent pas être négatives, il y'a un problème.")
        else:
            self.pmid = 1


    def proba_total(self, tree):
        """
        Met à jour la probabilité totale de se retrouver à ce nœud et aux nœuds adjacents.

        Args:
            tree: L'arbre trinomial utilisé pour le calcul.
        """
        if self.p_total > tree.seuil:
            self.n_mid.p_total += self.pmid * self.p_total
            self.n_up.p_total += self.pup * self.p_total
            self.n_down.p_total += self.pdown * self.p_total
        else:
            self.n_mid.p_total += self.pmid * self.p_total

    def build_block(self, n: 'Node', tree, have_div: bool):
        """
        Construit le bloc de nœuds adjacents pour le nœud courant.

        Args:
            n (Node): Le nœud suivant dans l'arbre.
            tree: L'arbre trinomial utilisé pour le calcul.
            have_div (bool): Indique si un dividende doit être pris en compte.
        """

        if self.p_total < tree.seuil and (self.up is None or self.down is None):

            self.n_mid = self.get_mid(n, have_div)
        else:
            # Créer le Node up si pas créé + branchement OU si il existe déjà fait simplement branchement

            self.n_mid = self.get_mid(n, have_div)

            if self.n_mid.up is None:
                self.n_up = Node(self.forward_up(have_div), self.i + 1, self.market, self.model)

                # Branchages Nmid / Nup
                self.n_up.down = self.n_mid
                self.n_mid.up = self.n_up
            else:
                self.n_up = self.n_mid.up

            # Créer le Node down si pas créé + branchement OU si il existe déjà fait simplement branchement
            if self.n_mid.down is None:
                self.n_down = Node(self.forward_down(have_div), self.i + 1, self.market, self.model)

                # Branchages Nmid / Ndown
                self.n_down.up = self.n_mid
                self.n_mid.down = self.n_down
            else:
                self.n_down = self.n_mid.down

        # CALCULE PROBA DE TRANSITION
        self.proba_transition(tree, have_div)   

        # CALCULE PROBA TOTAL de n_mid
        self.proba_total(tree)
