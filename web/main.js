// Import the WASM module
import init, { parse_tableau, run_rcd } from './pkg/ot_soft.js';

// Tiny example tableau data (loaded from examples/tiny/input.txt)
const TINY_EXAMPLE = `			*No Onset	*Coda	Max(t)	Dep(?)
			*NoOns	*Coda	Max	Dep
a	?a	1				1
	a	0	1
tat	ta	1		1
	tat	0		1
at	?a	1			1	1
	?at	0		1		1
	a	0	1		1
	at	0	1	1
`;

// Store current tableau and text globally
let currentTableau = null;
let currentTableauText = null;

async function run() {
    try {
        // Initialize the WASM module
        await init();

        // Hide loading message, show content
        document.getElementById('status').style.display = 'none';
        document.getElementById('content').style.display = 'block';

        // Set up file input handler
        document.getElementById('fileInput').addEventListener('change', async (event) => {
            const file = event.target.files[0];
            if (file) {
                const text = await file.text();
                parseAndDisplay(text);
            }
        });

        // Set up "Load Tiny Example" button
        document.getElementById('loadTinyButton').addEventListener('click', () => {
            parseAndDisplay(TINY_EXAMPLE);
        });

        // Set up "Run RCD" button
        document.getElementById('runRcdButton').addEventListener('click', () => {
            runRcdAnalysis();
        });

        console.log('OT-Soft WebAssembly module loaded successfully');
    } catch (err) {
        console.error('Failed to load WASM module:', err);
        document.getElementById('status').textContent = 'Error loading WebAssembly module. Check console for details.';
        document.getElementById('status').style.background = '#ffe8e8';
        document.getElementById('status').style.borderLeftColor = '#e74c3c';
    }
}

function parseAndDisplay(text) {
    try {
        const tableau = parse_tableau(text);
        currentTableau = tableau;
        currentTableauText = text;

        const html = formatTableauAsHTML(tableau);
        document.getElementById('tableauSection').style.display = 'block';
        document.getElementById('tableauOutput').innerHTML = html;

        // Show RCD section and hide previous results
        document.getElementById('rcdSection').style.display = 'block';
        document.getElementById('rcdOutput').style.display = 'none';
        document.getElementById('rcdOutput').innerHTML = '';
    } catch (err) {
        document.getElementById('tableauSection').style.display = 'block';
        document.getElementById('tableauOutput').textContent = 'Error parsing tableau:\n\n' + err;
        document.getElementById('rcdSection').style.display = 'none';
        console.error('Parse error:', err);
    }
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function formatTableauAsHTML(tableau) {
    let html = "<table class='tableau-table'>\n";

    const constraintCount = tableau.constraint_count();
    const formCount = tableau.form_count();

    // Row 1: Constraint full names (with empty cells for input/candidate/frequency)
    html += "  <tr class='header-row'>\n";
    html += "    <th></th>\n";  // Input column
    html += "    <th></th>\n";  // Candidate column
    html += "    <th></th>\n";  // Frequency column

    for (let i = 0; i < constraintCount; i++) {
        const constraint = tableau.get_constraint(i);
        html += `    <th>${escapeHtml(constraint.full_name)}</th>\n`;
    }
    html += "  </tr>\n";

    // Row 2: Column headers (Input, Candidate, Frequency, then constraint abbreviations)
    html += "  <tr class='subheader-row'>\n";
    html += "    <th>Input</th>\n";
    html += "    <th>Candidate</th>\n";
    html += "    <th>Freq</th>\n";

    for (let i = 0; i < constraintCount; i++) {
        const constraint = tableau.get_constraint(i);
        html += `    <th>${escapeHtml(constraint.abbrev)}</th>\n`;
    }
    html += "  </tr>\n";

    // Data rows
    for (let i = 0; i < formCount; i++) {
        const form = tableau.get_form(i);
        const candidateCount = form.candidate_count();

        for (let j = 0; j < candidateCount; j++) {
            html += "  <tr class='data-row'>\n";

            // Show input only on first candidate
            if (j === 0) {
                html += `    <td class='input-cell'>${escapeHtml(form.input)}</td>\n`;
            } else {
                html += "    <td class='input-cell'></td>\n";
            }

            // Candidate
            const candidate = form.get_candidate(j);
            html += `    <td class='candidate-cell'>${escapeHtml(candidate.form)}</td>\n`;

            // Frequency
            html += `    <td class='frequency-cell'>${candidate.frequency}</td>\n`;

            // Violations
            for (let k = 0; k < constraintCount; k++) {
                const violation = candidate.get_violation(k);
                const violationStr = (violation === 0 || violation === null) ? '' : violation.toString();
                html += `    <td class='violation-cell'>${violationStr}</td>\n`;
            }

            html += "  </tr>\n";
        }
    }

    html += "</table>\n";
    return html;
}

function runRcdAnalysis() {
    if (!currentTableauText) {
        alert('Please load a tableau file first');
        return;
    }

    try {
        const result = run_rcd(currentTableauText);
        displayRcdResults(result);
    } catch (err) {
        console.error('RCD error:', err);
        document.getElementById('rcdOutput').style.display = 'block';
        document.getElementById('rcdOutput').innerHTML = `
            <div class="rcd-status failure">
                Error running RCD: ${err}
            </div>
        `;
    }
}

function displayRcdResults(result) {
    const outputDiv = document.getElementById('rcdOutput');
    outputDiv.style.display = 'block';

    let html = '<div class="rcd-results">';

    // Status
    if (result.success()) {
        html += '<div class="rcd-status success">✓ A ranking was found that generates the correct outputs</div>';
    } else {
        html += '<div class="rcd-status failure">✗ Failed to find a valid ranking</div>';
    }

    // Group constraints by stratum
    const numStrata = result.num_strata();
    const constraintCount = currentTableau.constraint_count();

    const strata = [];
    for (let s = 1; s <= numStrata; s++) {
        strata.push([]);
    }

    for (let i = 0; i < constraintCount; i++) {
        const stratum = result.get_stratum(i);
        if (stratum && stratum >= 1 && stratum <= numStrata) {
            const constraint = currentTableau.get_constraint(i);
            strata[stratum - 1].push({
                abbrev: constraint.abbrev,
                fullName: constraint.full_name
            });
        }
    }

    // Display strata
    for (let s = 0; s < numStrata; s++) {
        html += '<div class="stratum">';
        html += `<div class="stratum-header">Stratum ${s + 1}</div>`;
        html += '<div class="constraint-list">';

        for (const constraint of strata[s]) {
            html += '<div class="constraint-item">';
            html += `<span class="abbrev">${escapeHtml(constraint.abbrev)}</span>`;
            html += `<span class="full-name">${escapeHtml(constraint.fullName)}</span>`;
            html += '</div>';
        }

        html += '</div>';
        html += '</div>';
    }

    html += '</div>';
    outputDiv.innerHTML = html;
}

run();
