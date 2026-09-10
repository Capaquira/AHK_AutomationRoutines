
function getConfirmedOperationsSummary() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const startRow = 64;
  const numRows = 95 - 64 + 1;
  const data = activeSheet.getRange(startRow, 2, numRows, 5).getValues(); 

  const serviceMap = new Map(); // clave = Reception/Loading/Stand-alone, valor = lista de servicios

  for (let i = 0; i < data.length; i++) {
    const descriptionRaw = String(data[i][0] || '').trim();
    const isConfirmed = String(data[i][3] || '').trim().toLowerCase();
    const specification = String(data[i][4] || '').trim();

    if (isConfirmed !== 'yes') continue;
    if (descriptionRaw.startsWith('-')) continue; // ignorar subservicios

    // Determinar si es servicio común o especial
    let description;
    if (
      descriptionRaw.toLowerCase().startsWith('reception') ||
      descriptionRaw.toLowerCase().startsWith('loading') ||
      descriptionRaw.toLowerCase().startsWith('stock inspection')
    ) {
      description = descriptionRaw; // mantener completo
    } else {
      description = descriptionRaw.split(' ')[0]; // solo primera palabra
    }

    if (specification) {
      specification.split(',').map(s => s.trim()).forEach(spec => {
        const key = spec; // Reception, Loading, Stand-alone
        if (!serviceMap.has(key)) {
          serviceMap.set(key, new Set());
        }
        serviceMap.get(key).add(description);
      });
    } else {
      if (!serviceMap.has(description)) {
        serviceMap.set(description, new Set());
      }
    }
  }

  const results = [];
  for (const [location, services] of serviceMap.entries()) {
    if (location.toLowerCase() === 'stand-alone') {
      services.forEach(s => results.push(s));
    } else if (services.size > 0) {
      const serviceList = [...services].join(' / ');
      results.push(`${location} (${serviceList})`);
    } else {
      results.push(location);
    }
  }

  // Evitar duplicados de Reception/Loading simples si ya existen con detalle
  const hasReceptionDetail = results.some(r => r.startsWith('Reception ('));
  const hasLoadingDetail = results.some(r => r.startsWith('Loading ('));

  const filtered = results.filter(r => {
    if (r === 'Reception' && hasReceptionDetail) return false;
    if (r === 'Loading' && hasLoadingDetail) return false;
    return true;
  });

  return formatNaturalList(filtered);
}

function formatNaturalList(items) {
  if (!items || items.length === 0) return '';
  if (items.length === 1) return items[0];
  if (items.length === 2) return `${items[0]} and ${items[1]}`;
  
  const lastItem = items[items.length - 1];
  const initialItems = items.slice(0, -1).join(', ');
  return `${initialItems} and ${lastItem}`;
}
