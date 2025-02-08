document.getElementById('tournamentForm').addEventListener('submit', function (event) {
    event.preventDefault();

    const playersInput = document.getElementById('players');
    const playersCount = parseInt(playersInput.value);

    const messageDiv = document.getElementById('message');
    const resultsDiv = document.getElementById('results');
    const exportButton = document.getElementById('exportExcel');

    // Clear previous messages or results
    messageDiv.innerHTML = '';
    messageDiv.classList.remove('success');
    resultsDiv.innerHTML = '';
    exportButton.style.display = 'none'; // Hide export button until tournament is generated

    // Validation
    if (isNaN(playersCount) || playersCount < 36) {
        messageDiv.textContent = "Il faut au moins 36 joueurs pour organiser 4 sessions complètes.";
        return;
    }

    // Initial setup
    const sessions = 4; // Number of sessions
    const players = Array.from({ length: playersCount }, (_, i) => `Joueur ${i + 1}`);
    const matchesPerSession = 6; // 6 matches per session
    const playersPerTeam = 3; // 3 players per team
    const teamsPerSession = matchesPerSession * 2; // 12 teams per session
    const teamsHistory = {}; // Keep track of teams to prevent duplicates

    function shuffle(array) {
        return array.sort(() => Math.random() - 0.5);
    }

    function generateSession(players, sessionNumber) {
        const sessionTeams = [];
        const availablePlayers = [...players];
        shuffle(availablePlayers);

        // Generate teams for the session
        for (let i = 0; i < teamsPerSession; i++) {
            let team = null;

            // Find a unique team (players not previously teamed together)
            do {
                team = {
                    tireur: availablePlayers.pop(),
                    milieu: availablePlayers.pop(),
                    pointeur: availablePlayers.pop()
                };
            } while (isDuplicateTeam(team));

            // Add team to session and history
            sessionTeams.push(team);
            recordTeamInHistory(team);
        }

        // Group teams into matches
        const matches = [];
        for (let i = 0; i < matchesPerSession; i++) {
            const match = {
                terrain: i + 1,
                equipe1: sessionTeams[i * 2],
                equipe2: sessionTeams[i * 2 + 1]
            };
            matches.push(match);
        }

        return matches;
    }

    function isDuplicateTeam(team) {
        const teamKey = teamKeyHash(team);
        return teamsHistory[teamKey] || false;
    }

    function recordTeamInHistory(team) {
        const teamKey = teamKeyHash(team);
        teamsHistory[teamKey] = true;
    }

    function teamKeyHash(team) {
        // Create a unique hash for the team by sorting the player names
        const players = [team.tireur, team.milieu, team.pointeur].sort();
        return players.join('-');
    }

    // Generate matches for all sessions
    const allSessionMatches = [];
    for (let sessionNumber = 1; sessionNumber <= sessions; sessionNumber++) {
        const sessionMatches = generateSession(players, sessionNumber);
        allSessionMatches.push({
            session: sessionNumber,
            matches: sessionMatches
        });
    }

    // Display results
    displayResults(allSessionMatches);
    exportButton.style.display = 'block'; // Show the export button when results are displayed

    // Set up export button
    exportButton.addEventListener('click', function () {
        exportToExcel(allSessionMatches);
    });
});

function displayResults(allSessionMatches) {
    const resultsDiv = document.getElementById('results');
    resultsDiv.innerHTML = `
        <h2>Programme du Tournoi</h2>
        ${allSessionMatches.map(session => `
            <div class="card">
                <h3>Session ${session.session}</h3>
                <div class="matches">
                    ${session.matches.map(match => `
                        <div class="match-card">
                            <h4>Terrain ${match.terrain}</h4>
                            <div class="teams">
                                <div class="team team-1">
                                    <h5>Équipe 1</h5>
                                    <ul>
                                        <li><strong>Tireur:</strong> ${match.equipe1.tireur}</li>
                                        <li><strong>Milieu:</strong> ${match.equipe1.milieu}</li>
                                        <li><strong>Pointeur:</strong> ${match.equipe1.pointeur}</li>
                                    </ul>
                                </div>
                                <div class="vs">
                                    <span>VS</span>
                                </div>
                                <div class="team team-2">
                                    <h5>Équipe 2</h5>
                                    <ul>
                                        <li><strong>Tireur:</strong> ${match.equipe2.tireur}</li>
                                        <li><strong>Milieu:</strong> ${match.equipe2.milieu}</li>
                                        <li><strong>Pointeur:</strong> ${match.equipe2.pointeur}</li>
                                    </ul>
                                </div>
                            </div>
                        </div>
                    `).join('')}
                </div>
            </div>
        `).join('')}
    `;
}


function exportToExcel(sessions) {
    // Construire les données pour le fichier Excel
    const workbook = XLSX.utils.book_new(); // Nouveau classeur

    sessions.forEach(session => {
        const sessionData = [];

        // En-têtes pour chaque session
        sessionData.push(['Terrain', 'Équipe 1 - Tireur', 'Équipe 1 - Milieu', 'Équipe 1 - Pointeur', 'Équipe 2 - Tireur', 'Équipe 2 - Milieu', 'Équipe 2 - Pointeur']);

        // Ajouter chaque match
        session.matches.forEach(match => {
            sessionData.push([
                `Terrain ${match.terrain}`,
                match.equipe1.tireur, match.equipe1.milieu, match.equipe1.pointeur,
                match.equipe2.tireur, match.equipe2.milieu, match.equipe2.pointeur
            ]);
        });

        // Convertir les données en une feuille Excel
        const worksheet = XLSX.utils.aoa_to_sheet(sessionData);
        XLSX.utils.book_append_sheet(workbook, worksheet, `Session ${session.session}`);
    });

    // Générer et télécharger le fichier Excel
    XLSX.writeFile(workbook, 'Tournoi_Petanque.xlsx');
}
