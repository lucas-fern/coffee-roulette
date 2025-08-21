document.addEventListener('DOMContentLoaded', function () {
    const generateBtn = document.getElementById('generateBtn');
    const clearBtn = document.getElementById('clearBtn');
    const csvInput = document.getElementById('csvInput');
    const groupSizeInput = document.getElementById('groupSize');
    const errorBox = document.getElementById('errorBox');
    const resultsContainer = document.getElementById('results');
    const downloadLink = document.getElementById('downloadCsv');
    let lastObjectUrl = null;

    // -------------------- UI helpers --------------------
    function showError(msg) {
        errorBox.innerText = msg;
        errorBox.style.display = 'block';
    }

    function clearError() {
        errorBox.innerText = '';
        errorBox.style.display = 'none';
    }

    // -------------------- Parsing --------------------
    function parseRoster(text) {
        const lines = text
            .split(/\r?\n/)
            .map((l) => l.trim())
            .filter((l) => l.length > 0);

        if (lines.length < 2) throw new Error('Need at least 2 people to form groups.');

        // Detect delimiter from the first non-empty line (tab or comma).
        const delimiter = lines[0].includes('\t') ? '\t' : ',';

        // Optional header: treat first row as header when it looks like "Name,Role".
        let startIdx = 0;
        const firstFields = lines[0].split(delimiter).map((s) => s.trim());
        if (
            firstFields.length >= 2 &&
            /^name$/i.test(firstFields[0]) &&
            /^role$/i.test(firstFields[1])
        ) {
            startIdx = 1;
        }

        const people = [];
        for (let i = startIdx; i < lines.length; i++) {
            const parts = lines[i].split(delimiter).map((s) => s.trim());
            const [name, role] = parts;
            if (!name || !role) {
                // +1 to show human-friendly (1-based) row number.
                throw new Error(`Line ${i + 1}: Name and Role must be non-empty.`);
            }
            people.push({ name, role });
        }
        return people;
    }

    // -------------------- Utilities --------------------
    function shuffleInPlace(arr) {
        for (let i = arr.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [arr[i], arr[j]] = [arr[j], arr[i]];
        }
        return arr;
    }

    function computeGroupSizes(n, desiredGroupSize) {
        if (!Number.isFinite(desiredGroupSize) || desiredGroupSize < 2) {
            throw new Error('Group Size must be at least 2.');
        }
        // Number of groups chosen to keep groups as even as possible.
        const g = Math.max(1, Math.round(n / desiredGroupSize));
        const base = Math.floor(n / g);
        const remainder = n - base * g;
        // First `remainder` groups get one extra member; diff between any two groups is <= 1.
        return Array.from({ length: g }, (_, i) => base + (i < remainder ? 1 : 0));
    }

    // Apportion role counts to groups using Hamilton's (largest remainder) method under group size constraints.
    // Returns an object: { role1: [c0, c1, ...], role2: [...], ... } where sum over roles per group equals group size
    // and sum over groups per role equals the total count of that role.
    function computeRoleTargets(sizes, countsByRole) {
        const roles = Object.keys(countsByRole);
        const g = sizes.length;
        const n = sizes.reduce((a, b) => a + b, 0);

        const targets = Object.fromEntries(roles.map((r) => [r, Array(g).fill(0)]));
        const remainders = Object.fromEntries(roles.map((r) => [r, Array(g).fill(0)]));
        const extras = Object.fromEntries(roles.map((r) => [r, 0]));
        const deficits = Array(g).fill(0);

        // Floors per cell and fractional remainders
        for (const r of roles) {
            const totalR = countsByRole[r];
            let sumFloors = 0;
            for (let i = 0; i < g; i++) {
                const q = (sizes[i] * totalR) / n; // ideal (possibly fractional) quota
                const f = Math.floor(q);
                targets[r][i] = f;
                remainders[r][i] = q - f;
                sumFloors += f;
            }
            extras[r] = totalR - sumFloors; // how many +1s for this role still need placing
        }

        // Group deficits given current floors across roles
        for (let i = 0; i < g; i++) {
            const filled = roles.reduce((acc, r) => acc + targets[r][i], 0);
            deficits[i] = sizes[i] - filled; // how many seats left in group i
        }

        // Global largest-remainders pass: consider all (group, role) cells ordered by remainder, but only
        // assign when the group still has capacity AND the role still has extras to place.
        const pairs = [];
        for (const r of roles) {
            for (let i = 0; i < g; i++) {
                pairs.push({ i, r, rem: remainders[r][i] });
            }
        }
        // Break ties fairly by shuffling before sorting by remainder.
        shuffleInPlace(pairs);
        pairs.sort((a, b) => b.rem - a.rem);

        for (const p of pairs) {
            if (extras[p.r] <= 0) continue;
            if (deficits[p.i] <= 0) continue;
            targets[p.r][p.i] += 1;
            extras[p.r] -= 1;
            deficits[p.i] -= 1;
        }

        // If any capacity remains (e.g., many exact-zero remainders), fill greedily with roles that still have extras,
        // preferring higher remainder for that group, then by larger remaining extras to keep spread even.
        let remainingDeficit = deficits.reduce((a, b) => a + b, 0);
        while (remainingDeficit > 0) {
            for (let i = 0; i < g && remainingDeficit > 0; i++) {
                while (deficits[i] > 0) {
                    let bestRole = null;
                    let bestRem = -1;
                    let bestExtras = -1;
                    for (const r of roles) {
                        if (extras[r] > 0) {
                            const rem = remainders[r][i];
                            if (rem > bestRem || (rem === bestRem && extras[r] > bestExtras)) {
                                bestRem = rem;
                                bestExtras = extras[r];
                                bestRole = r;
                            }
                        }
                    }
                    if (!bestRole) break; // No extras left anywhere (should not happen because totals must match)
                    targets[bestRole][i] += 1;
                    extras[bestRole] -= 1;
                    deficits[i] -= 1;
                    remainingDeficit -= 1;
                }
            }
            // Safety valve to avoid infinite loop in pathological cases
            const totalExtras = roles.reduce((acc, r) => acc + extras[r], 0);
            if (totalExtras === 0) break;
            remainingDeficit = deficits.reduce((a, b) => a + b, 0);
        }

        return targets;
    }

    // -------------------- Core allocation --------------------
    function allocateGroups(people, groupSize) {
        const n = people.length;
        const sizes = computeGroupSizes(n, groupSize);

        // Group people by role and shuffle within each role for random assignment.
        const byRole = {};
        for (const person of people) {
            if (!byRole[person.role]) byRole[person.role] = [];
            byRole[person.role].push(person);
        }
        for (const role in byRole) shuffleInPlace(byRole[role]);

        const roles = Object.keys(byRole);
        const counts = Object.fromEntries(roles.map((r) => [r, byRole[r].length]));

        const roleTargets = computeRoleTargets(sizes, counts);
        const g = sizes.length;
        const groups = Array.from({ length: g }, () => []);

        // Assign exactly according to targets, consuming from each role's pool so no person is used twice.
        for (const r of Object.keys(roleTargets)) {
            const pool = byRole[r] || [];
            let idx = 0;
            for (let gi = 0; gi < g; gi++) {
                const need = roleTargets[r][gi];
                for (let k = 0; k < need; k++) {
                    if (idx >= pool.length) {
                        throw new Error(`Internal error: insufficient people with role "${r}".`);
                    }
                    groups[gi].push(pool[idx]);
                    idx += 1;
                }
            }
            // Drop consumed people so they cannot be reused later
            byRole[r] = pool.slice(idx);
        }

        // In principle there should be no leftovers and all groups should be full.
        const leftovers = [];
        for (const r of Object.keys(byRole)) {
            if (byRole[r].length) leftovers.push(...byRole[r]);
        }
        shuffleInPlace(leftovers);
        for (let gi = 0; gi < g && leftovers.length > 0; gi++) {
            while (groups[gi].length < sizes[gi] && leftovers.length > 0) {
                groups[gi].push(leftovers.pop());
            }
        }

        // Sanity checks
        for (let gi = 0; gi < g; gi++) {
            if (groups[gi].length !== sizes[gi]) {
                throw new Error('Internal allocation error: group sizes do not match targets.');
            }
        }
        const assignedCount = groups.reduce((acc, gr) => acc + gr.length, 0);
        if (assignedCount !== n) {
            throw new Error('Internal allocation error: not all people were assigned exactly once.');
        }

        // Build annotated rows for CSV
        const annotated = [];
        for (let gi = 0; gi < g; gi++) {
            for (const person of groups[gi]) {
                annotated.push([person.name, person.role, gi + 1]);
            }
        }

        return [groups, annotated];
    }

    // -------------------- Rendering --------------------
    function renderResults(groups) {
        if (!groups.length) {
            resultsContainer.innerHTML = "<p class='muted'>No results yet.</p>";
            return;
        }

        const parts = groups
            .map((group, i) => {
                const rows = group
                    .map((person) => `<tr><td>${person.name}</td><td>${person.role}</td></tr>`)
                    .join('');
                return `<article class='panel' style='padding: .75rem; margin-bottom: .75rem;'>
                <h3 style='margin: .25rem 0 .5rem 0;'>Group ${i + 1} <small class='muted'>(n=${group.length})</small></h3>
                <div style='overflow:auto;'>
                    <table>
                        <thead>
                            <tr><th>Name</th><th>Role</th></tr>
                        </thead>
                        <tbody>${rows}</tbody>
                    </table>
                </div>
            </article>`;
            })
            .join('');

        resultsContainer.innerHTML = parts;
    }

    // -------------------- CSV download --------------------
    function buildDownload(annotated) {
        function csvEscape(value) {
            if (value == null) return '';
            const s = String(value);
            // Quote when needed and escape embedded quotes by doubling them per RFC 4180.
            if (/[",\n]/.test(s)) {
                return '"' + s.replace(/"/g, '""') + '"';
            }
            return s;
        }
        const header = 'Name,Role,Group';
        const body = annotated.map((row) => row.map(csvEscape).join(',')).join('\n');
        const csvContent = header + '\n' + body;
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });

        // Revoke previous URL (if any) to avoid leaks if user regenerates multiple times.
        if (lastObjectUrl) {
            try { URL.revokeObjectURL(lastObjectUrl); } catch { }
            lastObjectUrl = null;
        }

        const url = URL.createObjectURL(blob);
        lastObjectUrl = url;

        // Ensure the control behaves as a download link
        downloadLink.setAttribute('href', url);
        downloadLink.setAttribute('download', 'coffee_roulette_groups.csv');
        downloadLink.style.display = 'inline-flex';

        // Revoke after click to free memory.
        downloadLink.onclick = () => {
            setTimeout(() => {
                try { URL.revokeObjectURL(url); } catch { }
                if (lastObjectUrl === url) lastObjectUrl = null;
            }, 200);
        };
    }

    // -------------------- Event handlers --------------------
    function onGenerate() {
        clearError();
        resultsContainer.innerHTML = '';
        downloadLink.style.display = 'none';
        downloadLink.removeAttribute('href');

        let groupSize;
        try {
            groupSize = parseInt(groupSizeInput.value, 10);
            if (Number.isNaN(groupSize)) throw new Error();
        } catch {
            showError('Group Size must be an integer (e.g., 4).');
            return;
        }

        const text = csvInput.value || '';
        try {
            const roster = parseRoster(text);
            const [groups, annotated] = allocateGroups(roster, groupSize);
            renderResults(groups);
            buildDownload(annotated);
        } catch (e) {
            showError(e.message || String(e));
        }
    }

    function onClear() {
        clearError();
        csvInput.value = '';
        groupSizeInput.value = 4;
        resultsContainer.innerHTML = '';
        downloadLink.style.display = 'none';
        downloadLink.removeAttribute('href');
        if (lastObjectUrl) {
            try { URL.revokeObjectURL(lastObjectUrl); } catch { }
            lastObjectUrl = null;
        }
    }

    generateBtn.addEventListener('click', onGenerate);
    clearBtn.addEventListener('click', onClear);
});