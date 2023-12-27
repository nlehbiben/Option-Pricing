from Market import Market
from Option import Option
from Model import Model
from Node import Node
from Tree import Tree
from BlackScholes import BlackScholes
from ExcelInterface import ExcelInterface
class Convergence:

    def __init__(self, interface: ExcelInterface):
        self.interface = interface
        self.data = interface.read_data()
        self.is_pruned = self.data.get('is_pruned', 'Non') == 'Oui'
        self.print_arbre = False
    def run_trinomial(self) -> float:
        market = Market(**{k: self.data[k] for k in ['r', 'vol', 's0', 'div', 'div_date']})
        option = Option(**{k: self.data[k] for k in ['option_type', 'type', 'strike', 'maturity']})
        model = Model(pricing_date=self.data['pricing_date'], nbsteps=self.data['nbsteps'],
                      option=option, market=market)
        seuil = self.data['pruned_level'] if self.is_pruned else 0
        node = Node(self.data['s0'], i=0, market=market, model=model, p_total=1, seuil_enabled=self.is_pruned)
        tree = Tree(node, seuil=seuil, market=market, model=model)
    
        # Ne pas afficher l'arbre si print_tree est False
        if self.data['print_arbre']:
            tree.build_tree(output="S", print_tree=True, ws=self.interface.sht_arbre)
        else:
            tree.build_tree(output="S", print_tree=False)
    
        return node.price(option, tree)

    def run_black_scholes(self) -> dict:
        market = Market(**{k: self.data[k] for k in ['r', 'vol', 's0', 'div', 'div_date']})
        option = Option(**{k: self.data[k] for k in ['option_type', 'type', 'strike', 'maturity']})
        model = Model(pricing_date=self.data['pricing_date'], nbsteps=self.data['nbsteps'],
                      option=option, market=market)
        bs_pricer = BlackScholes(market, option, model)

        price = bs_pricer.price()
        greeks = bs_pricer.calculate_greeks()

        return {"Price": price, "Greeks": greeks}

    def convergence_nbsteps(self) -> list:
        max_steps = self.data['max_steps']
        convergence_results = []

        bs_result = self.run_black_scholes()  # Appeler run_black_scholes pour obtenir le résultat complet
        for nb_steps in range(1, max_steps + 1):
            self.data['nbsteps'] = nb_steps
            trinomial_price = self.run_trinomial()
            # Extraire le prix de Black & Scholes du résultat
            bs_price = bs_result["Price"]
            convergence_diff = (trinomial_price - bs_price) * nb_steps
            convergence_results.append([nb_steps, trinomial_price, bs_price, convergence_diff])
        return convergence_results

    def convergence_strike(self, nb_steps: int = 10) -> list:
        original_strike = int(self.data['strike'])  # Convertir le Strike en entier
        max_range = int(original_strike * 0.10)
        strike_range = list(range(original_strike - max_range, original_strike + max_range + 1))

        convergence_results = []
        previous_trinomial_price = None
        previous_bs_price = None

        for strike in strike_range:
            self.data['strike'] = strike
            self.data['nbsteps'] = nb_steps  # Fixer Nb_steps à 10
            trinomial_price = self.run_trinomial()
            bs_price = self.run_black_scholes()["Price"]

            if previous_trinomial_price is not None and previous_bs_price is not None:
                trinomial_slope = (trinomial_price - previous_trinomial_price) / (strike - (strike - 1))
                bs_slope = (bs_price - previous_bs_price) / (strike - (strike - 1))
            else:
                trinomial_slope = None
                bs_slope = None

            previous_trinomial_price = trinomial_price
            previous_bs_price = bs_price

            convergence_diff = trinomial_price - bs_price
            convergence_results.append([strike, trinomial_price, bs_price, convergence_diff, trinomial_slope, bs_slope])

        self.data['strike'] = original_strike  # Réinitialiser la valeur du Strike initial

        return convergence_results

