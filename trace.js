Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById('log-btn').addEventListener('click', handleButtonClick);
        document.getElementById('show-content-btn').addEventListener('click', showCellContent);
        document.getElementById('load-json-btn').addEventListener('click', loadJsonContent);
    }
    
    // Render the accounting equation header
    const mathHeader = document.getElementById('math-header');
    if (mathHeader && typeof katex !== 'undefined') {
        katex.render('\\text{Assets} = \\text{Liabilities} + \\text{Equity}', mathHeader, {
            throwOnError: false,
            displayMode: true
        });
    }
});

function handleButtonClick() {
    console.log('Button clicked!');
}

async function loadJsonContent() {
    try {
        const response = await fetch('multidim_dag_resolution/gross_profit_calc_graph.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        let data = await response.json();
        
        // Recursively resolve references
        data = await resolveReferences(data);

        const container = document.getElementById('json-display');
        container.innerHTML = ''; // Clear previous content
        container.className = 'box tree-view'; // Add tree-view class
        container.appendChild(createTreeView(data));
    } catch (error) {
        console.error('Error loading JSON:', error);
        document.getElementById('json-display').textContent = 'Error loading JSON: ' + error.message;
    }
}

async function resolveReferences(data, basePath = 'multidim_dag_resolution/') {
    if (data === null || typeof data !== 'object') {
        return data;
    }

    if (Array.isArray(data)) {
        return Promise.all(data.map(item => resolveReferences(item, basePath)));
    }

    if (data['$ref']) {
        const refPath = basePath + data['$ref'];
        const response = await fetch(refPath);
        if (!response.ok) {
            throw new Error(`Failed to load reference: ${refPath}`);
        }
        const refData = await response.json();
        return resolveReferences(refData, basePath);
    }

    const resolved = {};
    for (const key of Object.keys(data)) {
        resolved[key] = await resolveReferences(data[key], basePath);
    }
    return resolved;
}

function createTreeView(data, key = null) {
    // Handle primitive types (leaf nodes)
    if (data === null || typeof data !== 'object') {
        const span = document.createElement('span');
        if (key !== null) {
            const keySpan = document.createElement('span');
            keySpan.className = 'tree-key';
            keySpan.textContent = `"${key}": `;
            span.appendChild(keySpan);
        }

        const mathKeys = ['equation', 'term_expression', 'formula', 'set_predicate'];
        if (key && mathKeys.includes(key) && typeof data === 'string') {
            const mathSpan = document.createElement('span');
            mathSpan.style.margin = '0 5px';
            
            if (typeof katex !== 'undefined') {
                try {
                    katex.render(data, mathSpan, {
                        throwOnError: false
                    });
                } catch (e) {
                    console.warn('KaTeX error:', e);
                    mathSpan.textContent = JSON.stringify(data);
                }
            } else {
                mathSpan.textContent = JSON.stringify(data);
            }
            span.appendChild(mathSpan);
        } else {
            const valSpan = document.createElement('span');
            valSpan.className = `tree-${data === null ? 'null' : typeof data}`;
            valSpan.textContent = JSON.stringify(data);
            span.appendChild(valSpan);
        }
        return span;
    }

    // Handle Objects and Arrays (branch nodes)
    const container = document.createElement('div');
    
    const header = document.createElement('div');
    const toggle = document.createElement('span');
    toggle.className = 'tree-toggle';
    toggle.textContent = '[+]'; // Default collapsed
    
    const label = document.createElement('span');
    if (key !== null) {
        const keySpan = document.createElement('span');
        keySpan.className = 'tree-key';
        keySpan.textContent = `"${key}": `;
        label.appendChild(keySpan);
    }
    
    const isArray = Array.isArray(data);
    const keys = Object.keys(data);
    const count = keys.length;
    const openBrace = document.createElement('span');
    openBrace.textContent = isArray ? `[ ${count} items ]` : `{ ... }`;
    
    header.appendChild(toggle);
    header.appendChild(label);
    header.appendChild(openBrace);
    container.appendChild(header);

    const children = document.createElement('ul');
    children.style.display = 'none';
    
    keys.forEach((k, i) => {
        const li = document.createElement('li');
        // Pass key only if it's an object, for arrays we just show values usually, 
        // but to match JSON structure strictly we can omit key for array items or show index.
        // Standard JSON view: Object keys are shown, Array indices are usually implied.
        const childNode = createTreeView(data[k], isArray ? null : k);
        li.appendChild(childNode);
        if (i < keys.length - 1) {
            li.appendChild(document.createTextNode(','));
        }
        children.appendChild(li);
    });

    const closeBrace = document.createElement('div');
    closeBrace.textContent = isArray ? ']' : '}';
    closeBrace.style.paddingLeft = '24px'; // Indent closing brace
    closeBrace.style.display = 'none';

    container.appendChild(children);
    container.appendChild(closeBrace);

    // Toggle logic
    toggle.onclick = (e) => {
        e.stopPropagation();
        const isHidden = children.style.display === 'none';
        if (isHidden) {
            children.style.display = 'block';
            closeBrace.style.display = 'block';
            toggle.textContent = '[-]';
            openBrace.textContent = isArray ? '[' : '{';
        } else {
            children.style.display = 'none';
            closeBrace.style.display = 'none';
            toggle.textContent = '[+]';
            // Show summary when collapsed
            const count = keys.length;
            openBrace.textContent = isArray ? `[ ${count} items ]` : `{ ... }`;
        }
    };

    return container;
}

async function showCellContent() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load("values");
            await context.sync();

            const display = document.getElementById("display");
            if (range.values && range.values.length > 0) {
                display.innerText = range.values[0][0];
            } else {
                display.innerText = "No content";
            }
        });
    } catch (error) {
        console.error(error);
        document.getElementById("display").innerText = "Error: " + error.message;
    }
}
