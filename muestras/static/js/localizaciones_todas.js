const OPEN_NODES_STORAGE_KEY = 'geslab.recursos.openNodes';

function loadOpenNodes() {
    try {
        const raw = window.localStorage.getItem(OPEN_NODES_STORAGE_KEY);
        const parsed = raw ? JSON.parse(raw) : [];
        return new Set(Array.isArray(parsed) ? parsed : []);
    } catch (error) {
        return new Set();
    }
}

function saveOpenNodes(openNodes) {
    window.localStorage.setItem(OPEN_NODES_STORAGE_KEY, JSON.stringify(Array.from(openNodes)));
}

function getNodeKey(button) {
    return button ? button.dataset.nodeKey || '' : '';
}

function rememberNodeState(button, expanded) {
    const nodeKey = getNodeKey(button);
    if (!nodeKey) return;
    const openNodes = loadOpenNodes();
    if (expanded) {
        openNodes.add(nodeKey);
    } else {
        openNodes.delete(nodeKey);
    }
    saveOpenNodes(openNodes);
}

function setButtonArrow(button, expanded) {
    if (!button) return;
    if (!button.dataset.baseLabel) {
        button.dataset.baseLabel = button.innerHTML.trim();
    }
    button.innerHTML = button.dataset.baseLabel + ' ' + (expanded ? '&#9660;' : '&#9654;');
    button.dataset.expanded = expanded ? 'true' : 'false';
}

function setExpandedState(button, element, expanded, remember) {
    if (!button || !element) return;
    element.style.display = expanded ? 'block' : 'none';
    setButtonArrow(button, expanded);
    if (remember) {
        rememberNodeState(button, expanded);
    }
}

function desplegar(button, className, remember = true) {
    const element = button.parentElement.querySelector(':scope > .' + className);
    if (!element) return;
    const expanded = element.style.display !== 'block';
    setExpandedState(button, element, expanded, remember);
}

function buildBoxGrid(ul, data) {
    ul.innerHTML = '';
    if (!data.subposiciones || data.subposiciones.length === 0) {
        ul.innerHTML = '<li style="color:#888; font-style:italic; padding:6px 0;">Sin subposiciones</li>';
        ul.dataset.loaded = 'true';
        return;
    }

    ul.classList.add('box-visual-list');
    const grouped = {};
    data.subposiciones.forEach(function(sub) {
        const rowKey = sub.fila || '-';
        if (!grouped[rowKey]) grouped[rowKey] = [];
        grouped[rowKey].push(sub);
    });

    function columnSortValue(value) {
        const numeric = Number(value);
        return Number.isNaN(numeric) ? String(value) : numeric;
    }

    Object.keys(grouped)
        .sort()
        .forEach(function(rowKey) {
            grouped[rowKey].sort(function(a, b) {
                const aValue = columnSortValue(a.columna || '');
                const bValue = columnSortValue(b.columna || '');
                if (typeof aValue === 'number' && typeof bValue === 'number') return aValue - bValue;
                return String(aValue).localeCompare(String(bValue), undefined, { numeric: true });
            });

            const rowLi = document.createElement('li');
            rowLi.className = 'box-visual-row';

            const rowLabel = document.createElement('div');
            rowLabel.className = 'box-row-label';
            rowLabel.textContent = rowKey;
            rowLi.appendChild(rowLabel);

            const cells = document.createElement('div');
            cells.className = 'box-row-cells';

            grouped[rowKey].forEach(function(sub) {
                const cell = document.createElement('label');
                cell.className = 'box-cell-visual ' + (sub.vacia ? 'box-cell-empty' : 'box-cell-filled');

                let contenidoMuestra = '-';
                if (!sub.vacia && sub.muestra_nom_lab) {
                    contenidoMuestra = sub.muestra_nom_lab;
                }

                let mainContent = '<span class="box-cell-main">' + contenidoMuestra + '</span>';
                if (!sub.vacia && sub.muestra_detalle_url) {
                    mainContent =
                        '<a href="' + sub.muestra_detalle_url + '" class="box-cell-main box-cell-link" title="Abrir detalle de la muestra">' +
                        contenidoMuestra +
                        '</a>';
                }

                cell.innerHTML =
                    '<input type="checkbox" name="subposicion" value="' + sub.id + '" class="box-cell-checkbox" onchange="actualizarEstadoBotonEliminar()">' +
                    '<span class="box-cell-corner">' + (sub.columna || '-') + '</span>' +
                    mainContent;

                cells.appendChild(cell);
            });

            rowLi.appendChild(cells);
            ul.appendChild(rowLi);
        });

    ul.dataset.loaded = 'true';
}

function desplegarCaja(button, cajaId, remember = true) {
    const ul = button.parentElement.querySelector('ul.muestras[data-caja-id="' + cajaId + '"]');
    if (!ul) return;

    if (ul.style.display === 'block') {
        setExpandedState(button, ul, false, remember);
        return;
    }

    setExpandedState(button, ul, true, remember);

    if (ul.dataset.loaded === 'true') return;

    ul.innerHTML = '<li style="color:#888; font-style:italic; padding:6px 0;"><span class="spinner-sm"></span> Cargando subposiciones...</li>';

    fetch('/api/get_subposiciones_por_caja_tree/?caja_id=' + cajaId)
        .then(function(response) { return response.json(); })
        .then(function(data) {
            buildBoxGrid(ul, data);
        })
        .catch(function() {
            ul.innerHTML = '<li style="color:#c00; padding:6px 0;">Error al cargar subposiciones</li>';
        });
}

function restoreOpenNodes() {
    const openNodes = loadOpenNodes();
    if (!openNodes.size) return;

    document.querySelectorAll('.dropbtn[data-node-key]').forEach(function(button) {
        const nodeKey = getNodeKey(button);
        if (!openNodes.has(nodeKey)) return;

        if (button.dataset.cajaId) {
            desplegarCaja(button, button.dataset.cajaId, false);
            return;
        }

        const element = button.parentElement.querySelector(':scope > ul');
        if (!element) return;
        setExpandedState(button, element, true, false);
    });
}

document.addEventListener('DOMContentLoaded', function() {
    document.querySelectorAll(
        '.estantes, .posicion_estante, .racks, .bandejas, .posicion_caja_rack, .cajas, .muestras'
    ).forEach(function(el) {
        el.style.display = 'none';
    });

    document.querySelectorAll('.dropbtn').forEach(function(btn) {
        setButtonArrow(btn, false);
    });

    restoreOpenNodes();
});
