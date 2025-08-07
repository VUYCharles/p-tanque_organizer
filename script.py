import random
from itertools import combinations
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

class Player:
    def __init__(self, id, name, role):
        self.id = id
        self.name = name
        self.role = role  # 'tireur', 'pointeur', 'milieu'

    def __repr__(self):
        return f"{self.name} ({self.role})"

class Team:
    def __init__(self, shooter, pointer, middle, team_id):
        self.shooter = shooter
        self.pointer = pointer
        self.middle = middle
        self.id = team_id

    def __repr__(self):
        return f"Équipe {self.id}: [T: {self.shooter.name}, P: {self.pointer.name}, M: {self.middle.name}]"

    def full_info(self):
        return (f"Équipe {self.id}\n"
                f"Tireur: {self.shooter.name}\n"
                f"Pointeur: {self.pointer.name}\n"
                f"Milieu: {self.middle.name}")

class Match:
    def __init__(self, team1, team2):
        self.team1 = team1
        self.team2 = team2

    def __repr__(self):
        return f"{self.team1} vs {self.team2}"

class Tournament:
    def __init__(self, players, num_courts):
        self.players = players
        self.num_courts = num_courts
        self.num_teams = len(players) // 3
        self.rounds = []
        self.teams_history = []
        self.shooter_encounters = set()
        self.team_compositions = set()

    def generate_teams(self):
        shooters = [p for p in self.players if p.role == 'tireur']
        pointers = [p for p in self.players if p.role == 'pointeur']
        middles = [p for p in self.players if p.role == 'milieu']

        random.shuffle(shooters)
        random.shuffle(pointers)
        random.shuffle(middles)

        return [Team(shooters[i], pointers[i], middles[i], i+1) for i in range(self.num_teams)]

    def is_valid_team_composition(self, teams):
        current_compositions = set()
        for team in teams:
            composition = (team.shooter.id, team.pointer.id, team.middle.id)
            if composition in self.team_compositions:
                return False
            current_compositions.add(composition)
        self.team_compositions.update(current_compositions)
        return True

    def generate_matches(self, teams):
        possible_pairs = list(combinations(teams, 2))
        random.shuffle(possible_pairs)

        matches = []
        used_teams = set()
        new_shooter_pairs = set()

        for team1, team2 in possible_pairs:
            if team1.id in used_teams or team2.id in used_teams:
                continue

            shooter_pair = frozenset({team1.shooter.id, team2.shooter.id})
            if shooter_pair not in self.shooter_encounters:
                matches.append(Match(team1, team2))
                used_teams.add(team1.id)
                used_teams.add(team2.id)
                new_shooter_pairs.add(shooter_pair)
                if len(matches) == self.num_courts:
                    break

        remaining_teams = [t for t in teams if t.id not in used_teams]
        while len(remaining_teams) >= 2 and len(matches) < self.num_courts:
            team1 = remaining_teams.pop()
            team2 = remaining_teams.pop()
            shooter_pair = frozenset({team1.shooter.id, team2.shooter.id})
            if shooter_pair not in self.shooter_encounters:
                matches.append(Match(team1, team2))
                new_shooter_pairs.add(shooter_pair)

        self.shooter_encounters.update(new_shooter_pairs)
        return matches if len(matches) == self.num_courts else None

    def play_round(self, round_num):
        max_attempts = 1000
        for _ in range(max_attempts):
            teams = self.generate_teams()
            if not self.is_valid_team_composition(teams):
                continue

            matches = self.generate_matches(teams)
            if matches:
                self.teams_history.append(teams)
                self.rounds.append(matches)
                return True
        return False

    def run_tournament(self):
        num_rounds = 6  # Nombre de rondes standard
        if self.num_teams == 8:  # 24 joueurs
            num_rounds = 4
        elif self.num_teams == 10:  # 30 joueurs
            num_rounds = 5
        return all(self.play_round(i+1) for i in range(num_rounds))

    def display_tournament(self):
        for i, (round_matches, teams) in enumerate(zip(self.rounds, self.teams_history), 1):
            print(f"\n=== Ronde {i} ===")
            print("\nComposition des équipes:")
            for team in teams:
                print(f"  {team}")
            print("\nMatches:")
            for j, match in enumerate(round_matches, 1):
                print(f"  Match {j}: {match}")

def generate_players(num_players):
    num_teams = num_players // 3
    roles = ['tireur']*num_teams + ['pointeur']*num_teams + ['milieu']*num_teams
    return [Player(i+1, f"Joueur_{i+1:02d}", roles[i]) for i in range(num_players)]

def validate_constraints(tournament):
    print("\n=== Validation des contraintes ===")

    expected_rounds = 4 if tournament.num_teams == 8 else (5 if tournament.num_teams == 10 else 6)
    assert len(tournament.rounds) == expected_rounds, f"Devrait avoir {expected_rounds} rondes"
    print(f"✓ {expected_rounds} rondes organisées")

    for round_teams in tournament.teams_history:
        assert len(round_teams) == tournament.num_teams, f"Devrait avoir {tournament.num_teams} équipes par ronde"
    print(f"✓ {tournament.num_teams} équipes par ronde")

    all_compositions = set()
    for round_teams in tournament.teams_history:
        for team in round_teams:
            composition = (team.shooter.id, team.pointer.id, team.middle.id)
            assert composition not in all_compositions, "Composition d'équipe en double"
            all_compositions.add(composition)
    print("✓ Compositions d'équipes uniques pour toutes les rondes")

    shooter_pairs = set()
    for round_matches in tournament.rounds:
        for match in round_matches:
            pair = frozenset({match.team1.shooter.id, match.team2.shooter.id})
            assert pair not in shooter_pairs, "Tireurs se rencontrant plusieurs fois"
            shooter_pairs.add(pair)
    print("✓ Les tireurs ne se rencontrent qu'une seule fois")

    for round_teams in tournament.teams_history:
        player_ids = set()
        for team in round_teams:
            player_ids.add(team.shooter.id)
            player_ids.add(team.pointer.id)
            player_ids.add(team.middle.id)
        assert len(player_ids) == len(tournament.players), "Tous les joueurs doivent participer à chaque ronde"
    print("✓ Tous les joueurs participent à chaque ronde")

    for round_matches in tournament.rounds:
        assert len(round_matches) == tournament.num_courts, f"Devrait avoir {tournament.num_courts} matches par ronde"
    print(f"✓ {tournament.num_courts} matches par ronde")

    print("\n✓✓✓ Toutes les contraintes sont validées avec succès ! ✓✓✓")

def generate_excel(tournament, filename="tournoi_petanque.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Plan du tournoi"

    # Styles
    title_font = Font(bold=True, size=16, color="1F497D")
    round_header_font = Font(bold=True, size=12, color="FFFFFF")
    team_header_font = Font(bold=True, color="FFFFFF")
    match_header_font = Font(bold=True, color="FFFFFF")

    round_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    team_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    match_fill = PatternFill(start_color="76933C", end_color="76933C", fill_type="solid")

    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))

    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Titre principal
    ws.merge_cells('A1:E1')
    title_cell = ws['A1']
    title_cell.value = "TOURNOI DE PÉTANQUE - PLAN COMPLET"
    title_cell.font = title_font
    title_cell.alignment = center_alignment

    # Configuration des colonnes
    columns = [
        ('A', 'Ronde', 8),
        ('B', 'Équipe', 12),
        ('C', 'Tireur', 20),
        ('D', 'Pointeur', 20),
        ('E', 'Milieu', 20)
    ]

    for col, header, width in columns:
        ws.column_dimensions[col].width = width
        if header:
            cell = ws.cell(row=3, column=ord(col)-64, value=header)
            cell.font = team_header_font
            cell.fill = team_fill
            cell.alignment = center_alignment
            cell.border = border

    current_row = 4

    # Ajout des données des rondes
    for round_num, (teams, matches) in enumerate(zip(tournament.teams_history, tournament.rounds), 1):
        # En-tête de ronde
        ws.merge_cells(f'A{current_row}:E{current_row}')
        round_cell = ws[f'A{current_row}']
        round_cell.value = f"RONDE {round_num}"
        round_cell.font = round_header_font
        round_cell.fill = round_fill
        round_cell.alignment = center_alignment
        current_row += 1

        # Composition des équipes
        for team in teams:
            ws[f'A{current_row}'] = f"R{round_num}"
            ws[f'B{current_row}'] = f"Équipe {team.id}"
            ws[f'C{current_row}'] = team.shooter.name
            ws[f'D{current_row}'] = team.pointer.name
            ws[f'E{current_row}'] = team.middle.name

            for col in ['A', 'B', 'C', 'D', 'E']:
                cell = ws[f'{col}{current_row}']
                cell.border = border
                cell.alignment = left_alignment if col != 'A' else center_alignment

            current_row += 1

        # Section des matches
        ws.merge_cells(f'A{current_row}:E{current_row}')
        match_header_cell = ws[f'A{current_row}']
        match_header_cell.value = f"MATCHES DE LA RONDE {round_num}"
        match_header_cell.font = round_header_font
        match_header_cell.fill = round_fill
        match_header_cell.alignment = center_alignment
        current_row += 1

        # En-têtes des matches
        match_headers = ['Match', 'Équipe 1', '', 'Équipe 2', '']
        for col, header in zip(['A', 'B', 'C', 'D', 'E'], match_headers):
            if header:
                cell = ws[f'{col}{current_row}']
                cell.value = header
                cell.font = match_header_font
                cell.fill = match_fill
                cell.alignment = center_alignment
                cell.border = border
        current_row += 1

        # Détails des matches
        for match_num, match in enumerate(matches, 1):
            ws[f'A{current_row}'] = f"M{match_num}"
            ws[f'B{current_row}'] = match.team1.full_info()
            ws[f'D{current_row}'] = match.team2.full_info()
            ws[f'C{current_row}'] = "contre"

            for col in ['A', 'B', 'C', 'D', 'E']:
                cell = ws[f'{col}{current_row}']
                cell.border = border
                cell.alignment = center_alignment if col in ['A', 'C'] else left_alignment

            current_row += 1

        # Ajout d'espace entre les rondes
        current_row += 2

    # Feuille des rencontres de tireurs
    ws_encounters = wb.create_sheet("Rencontres de tireurs")

    # Titre
    ws_encounters.merge_cells('A1:C1')
    title_cell = ws_encounters['A1']
    title_cell.value = "RENCONTRES ENTRE TIREURS"
    title_cell.font = title_font
    title_cell.alignment = center_alignment

    # En-têtes
    headers = ['Ronde', 'Tireur 1', 'Tireur 2']
    for col, header in enumerate(headers, 1):
        cell = ws_encounters.cell(row=3, column=col, value=header)
        cell.font = team_header_font
        cell.fill = team_fill
        cell.alignment = center_alignment
        cell.border = border

    # Données
    row = 4
    for round_num, matches in enumerate(tournament.rounds, 1):
        for match in matches:
            ws_encounters.cell(row=row, column=1, value=f"Ronde {round_num}").border = border
            ws_encounters.cell(row=row, column=2, value=match.team1.shooter.name).border = border
            ws_encounters.cell(row=row, column=3, value=match.team2.shooter.name).border = border
            row += 1

    # Ajustement des largeurs de colonnes
    for col in ['A', 'B', 'C']:
        ws_encounters.column_dimensions[col].width = 25

    wb.save(filename)
    print(f"\nFichier Excel généré avec succès : {filename}")

def main():
    print("Organisation d'un tournoi de pétanque")

    while True:
        try:
            num_players = int(input("Nombre de joueurs (24, 30 ou 36) : "))
            if num_players in {24, 30, 36}:
                break
            print("Veuillez entrer 24, 30 ou 36")
        except ValueError:
            print("Veuillez entrer un nombre valide")

    num_courts = num_players // 6
    print(f"\nConfiguration :")
    print(f"- Nombre de joueurs : {num_players}")
    print(f"- Nombre d'équipes : {num_players // 3}")
    print(f"- Nombre de terrains disponibles : {num_courts}")

    players = generate_players(num_players)
    tournament = Tournament(players, num_courts)

    if tournament.run_tournament():
        tournament.display_tournament()
        validate_constraints(tournament)
        generate_excel(tournament)
    else:
        print("Impossible d'organiser le tournoi avec les contraintes données.")

if __name__ == "__main__":
    main()
