# Organisateur de Tournoi de Pétanque

Un outil Python pour organiser des tournois de pétanque avec des équipes équilibrées et des contraintes spécifiques.

## Fonctionnalités

- Génération automatique de tournois avec différentes configurations (24, 30 ou 36 joueurs)
- Création d'équipes équilibrées (tireur, pointeur, milieu)
- Respect des contraintes :
  - Unicité des compositions d'équipe
  - Rencontres uniques entre tireurs
  - Participation de tous les joueurs à chaque ronde
- Export vers Excel avec mise en forme professionnelle
- Interface graphique intuitive

## Prérequis

- Python 3.6 ou supérieur
- Packages Python :
  - openpyxl
  - tkinter (généralement inclus avec Python)

## Installation

1. Clonez le dépôt ou téléchargez les fichiers :
  `git clone https://github.com/votre-utilisateur/organisateur-petanque.git`
  `cd organisateur-petanque`


2. Installez les dépendances :
   `pip install openpyxl`

3. Si tkinter n'est pas installé (sous Linux généralement) :
   `sudo apt-get install python3-tk`
   
## Utilisation

1. Lancez l'application :
   `python petanque_tournament.py`

2. Dans l'interface graphique :
- Sélectionnez le nombre de joueurs (24, 30 ou 36)
- Cliquez sur "Générer le Tournoi"
- Consultez les résultats dans les onglets :
  - Composition des Équipes
  - Matches
  - Rencontres de Tireurs
- Exportez vers Excel avec le bouton dédié

## Format du tournoi

- **24 joueurs** (8 équipes) : 4 rondes, 4 terrains
- **30 joueurs** (10 équipes) : 5 rondes, 5 terrains
- **36 joueurs** (12 équipes) : 6 rondes, 6 terrains

## Personnalisation

Pour modifier les joueurs (remplacer les joueurs générés automatiquement), éditez la méthode `generate_players` dans la classe `PetanqueTournamentApp`.

## Licence

Ce projet est sous licence MIT. Voir le fichier LICENSE pour plus de détails.

## Contribution

Les contributions sont les bienvenues ! Ouvrez une issue ou soumettez une pull request.

---

*Développé avec passion pour la pétanque*
