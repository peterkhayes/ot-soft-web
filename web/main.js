// Import the WASM module
import init, { parse_tableau } from './pkg/ot_soft.js';

// Tiny example tableau data
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
        const html = formatTableauAsHTML(tableau);
        document.getElementById('outputSection').style.display = 'block';
        document.getElementById('output').innerHTML = html;
    } catch (err) {
        document.getElementById('outputSection').style.display = 'block';
        document.getElementById('output').textContent = 'Error parsing tableau:\n\n' + err;
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

run();
