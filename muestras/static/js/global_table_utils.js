function initSortableTable(selector, options = {}) {
    const table = document.querySelector(selector);
    if (!table) return;
    
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    
    // Header sorting function
    const headers = thead.querySelectorAll('th');
    const sortStates = {};
    const excludeColumns = options.excludeColumns || [];
    
    headers.forEach((header, index) => {
        if (excludeColumns.includes(index)) return; // skip specified columns
        
        header.style.cursor = 'pointer';
        header.style.userSelect = 'none';
        sortStates[index] = 0;
        header.classList.add('sorting');
        
        header.addEventListener('click', function() {
          if (sortStates[index] === 0) {
            sortStates[index] = 1;
          } else {
            sortStates[index] = sortStates[index] === 1 ? 2 : 1;
          }
          
          headers.forEach((h, i) => {
            if (i !== index) {
              h.classList.remove('sorting_asc', 'sorting_desc');
              if (!excludeColumns.includes(i)) {
                  h.classList.add('sorting');
              }
              sortStates[i] = 0;
            }
          });
          
          // Apply classes based on sort state
          header.classList.remove('sorting', 'sorting_asc', 'sorting_desc');
          if (sortStates[index] === 1) {
            header.classList.add('sorting_asc');
          } else if (sortStates[index] === 2) {
            header.classList.add('sorting_desc');
          } else {
            header.classList.add('sorting');
          }
          
          sortTable(index, sortStates[index]);
        });
    });
    
    function sortTable(colIndex, direction) {
        const sorted = [...rows].sort((a, b) => {
            const cellA = a.cells[colIndex];
            const cellB = b.cells[colIndex];
            const aSort = cellA.getAttribute('data-sort');
            const bSort = cellB.getAttribute('data-sort');
            // If cells have data-sort (dates), use that value
            if (aSort !== null && bSort !== null) {
                const aKey = aSort || '';
                const bKey = bSort || '';
                const cmp = aKey.localeCompare(bKey);
                return direction === 1 ? cmp : -cmp;
            }
            const aVal = cellA.textContent.trim();
            const bVal = cellB.textContent.trim();
            
            const aNum = parseFloat(aVal);
            const bNum = parseFloat(bVal);
            
            if (!isNaN(aNum) && !isNaN(bNum)) {
                return direction === 1 ? aNum - bNum : bNum - aNum;
            }
            
            return direction === 1 ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
        });
        
        tbody.innerHTML = '';
        sorted.forEach(row => tbody.appendChild(row));
    }
}
