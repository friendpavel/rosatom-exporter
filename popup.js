function updateProgress(progressElement, percent, text) {
  if (!progressElement) return;
  const fill = progressElement.querySelector('.progress-fill');
  const textEl = progressElement.querySelector('.progress-text');
  if (fill) {
    fill.style.width = `${percent}%`;
    fill.textContent = `${percent}%`;
  }
  if (textEl) textEl.textContent = text;
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function sanitizeFilename(name) {
  if (!name || name === 'Закупка') return 'Закупка';
  return name.replace(/[<>:"/\\|?*]/g, '_').substring(0, 200);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function createExcelFile(tableData) {
  try {
    if (!tableData || !tableData.rows || tableData.rows.length === 0) {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([['Нет данных'], ['Таблица "Детализированная потребность" не найдена']]);
      XLSX.utils.book_append_sheet(wb, ws, 'Потребность');
      return new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], 
        { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }
    
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([tableData.headers, ...tableData.rows]);
    
    ws['!cols'] = tableData.headers.map((h, i) => {
      const maxLen = Math.max(
        h.length,
        ...tableData.rows.map(row => String(row[i] || '').length)
      );
      return { wch: Math.min(maxLen + 2, 50) };
    });
    
    XLSX.utils.book_append_sheet(wb, ws, 'Потребность');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    return new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  } catch (e) {
    console.error('Excel ошибка:', e);
    return null;
  }
}

function createXMLFile(data) {
  function escapeXml(unsafe) {
    if (!unsafe) return '';
    return unsafe
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
  
  function makeSafeTagName(name) {
    return name
      .replace(/[^\wа-яё0-9]/gi, '_')
      .replace(/_{2,}/g, '_')
      .replace(/^_|_$/g, '');
  }
  
  let xml = '<?xml version="1.0" encoding="UTF-8"?>\n';
  xml += '<!-- Документ создан расширением Росатом для выгрузки данных -->\n';
  xml += '<ОсновныеДанные>\n';
  xml += `  <ДатаЭкспорта>${new Date().toLocaleString('ru-RU')}</ДатаЭкспорта>\n`;
  xml += `  <URL>${escapeXml(data.url || '')}</URL>\n\n`;
  
  for (const [key, value] of Object.entries(data)) {
    if (key === 'url') continue;
    if (key === 'ОКПД2_массив') continue;
    if (key === 'ОКВЭД2_массив') continue;
    if (key === 'Заказчики_массив') continue;
    if (typeof value === 'object') continue;
    if (key === 'downloadedCount') continue;
    
    if (value !== undefined && value !== null && value !== '') {
      const tagName = makeSafeTagName(key);
      const escapedValue = escapeXml(String(value).trim());
      xml += `  <${tagName}>${escapedValue}</${tagName}>\n`;
    }
  }
  
  if (data.Заказчики_массив && Array.isArray(data.Заказчики_массив) && data.Заказчики_массив.length > 0) {
    xml += `  <Заказчики_список>\n`;
    for (const customer of data.Заказчики_массив) {
      xml += `    <Заказчик>\n`;
      xml += `      <Наименование>${escapeXml(customer.name)}</Наименование>\n`;
      xml += `      <Главный>${customer.isMain ? 'Да' : 'Нет'}</Главный>\n`;
      xml += `    </Заказчик>\n`;
    }
    xml += `  </Заказчики_список>\n`;
  }
  
  if (data.ОКПД2_массив && Array.isArray(data.ОКПД2_массив) && data.ОКПД2_массив.length > 0) {
    xml += `  <ОКПД2_список>\n`;
    for (const item of data.ОКПД2_массив) {
      xml += `    <ОКПД2_запись>\n`;
      xml += `      <Код>${escapeXml(item.code)}</Код>\n`;
      xml += `      <Наименование>${escapeXml(item.name)}</Наименование>\n`;
      xml += `    </ОКПД2_запись>\n`;
    }
    xml += `  </ОКПД2_список>\n`;
  }
  
  if (data.ОКВЭД2_массив && Array.isArray(data.ОКВЭД2_массив) && data.ОКВЭД2_массив.length > 0) {
    xml += `  <ОКВЭД2_список>\n`;
    for (const item of data.ОКВЭД2_массив) {
      xml += `    <ОКВЭД2_запись>\n`;
      xml += `      <Код>${escapeXml(item.code)}</Код>\n`;
      xml += `      <Наименование>${escapeXml(item.name)}</Наименование>\n`;
      xml += `    </ОКВЭД2_запись>\n`;
    }
    xml += `  </ОКВЭД2_список>\n`;
  }
  
  xml += '</ОсновныеДанные>';
  return new Blob([xml], { type: 'application/xml;charset=utf-8' });
}

function switchToTab(tabName) {
  try {
    const tabs = document.querySelectorAll('.tabs-item');
    for (const tab of tabs) {
      if (tab.textContent && tab.textContent.includes(tabName)) {
        tab.click();
        return true;
      }
    }
    return false;
  } catch(e) {
    return false;
  }
}

function extractAllMainData() {
  const data = {
    url: window.location.href
  };
  
  try {
    const nameEl = document.querySelector('.text[style*="font-size: 30px"]');
    if (nameEl) data['Предмет_закупки'] = nameEl.textContent.trim();
    
    const numberEl = document.querySelector('.text[style*="font-size: 20px"]');
    if (numberEl) data['Номер_закупки'] = numberEl.textContent.trim();
    
    const allItems = document.querySelectorAll('.af-item');
    for (const item of allItems) {
      const titleEl = item.querySelector('.af-item__title');
      if (!titleEl) continue;
      
      let title = titleEl.textContent.trim();
      if (!title || title.length < 2) continue;
      
      let value = '';
      const valueSelectors = [
        '.af-core-label .text',
        '.af-item__content .text',
        '.af-core-label',
        '.text:not(.af-item__title)'
      ];
      
      for (const selector of valueSelectors) {
        const el = item.querySelector(selector);
        if (el) {
          value = el.textContent.trim();
          if (value && value !== title) break;
        }
      }
      
      if (value.startsWith(title)) {
        value = value.substring(title.length).trim();
      }
      
      if (!value || value === '') {
        value = '-';
      }
      
      data[title] = value;
    }
    
    const allCheckboxes = document.querySelectorAll('.af-checkbox, .af-checkbox-v2');
    for (const checkbox of allCheckboxes) {
      let parentItem = checkbox.closest('.af-item');
      if (!parentItem) continue;
      
      const titleEl = parentItem.querySelector('.af-item__title');
      if (!titleEl) continue;
      
      const label = titleEl.textContent.trim();
      if (!label || label.length < 2) continue;
      
      const isChecked = checkbox.querySelector('input[type="checkbox"]')?.checked;
      data[label] = isChecked ? 'Да' : 'Нет';
    }
    
    const customers = [];
    let customersTable = null;
    const allTables = document.querySelectorAll('table');
    
    for (const table of allTables) {
      const prevElement = table.previousElementSibling;
      const headerText = prevElement?.textContent?.trim() || '';
      if (headerText.includes('Заказчики') || headerText.includes('Заказчик')) {
        customersTable = table;
        break;
      }
      
      const th = table.querySelector('th');
      if (th && th.textContent.includes('Заказчик')) {
        customersTable = table;
        break;
      }
    }
    
    if (!customersTable) {
      customersTable = document.querySelector('[id*="Заказчики"] table, [class*="Заказчики"] table');
    }
    
    if (!customersTable) {
      const sections = document.querySelectorAll('.af-item');
      for (const section of sections) {
        const titleEl = section.querySelector('.af-item__title');
        if (titleEl && titleEl.textContent.trim() === 'Заказчики') {
          customersTable = section.querySelector('table');
          if (customersTable) break;
        }
      }
    }
    
    if (customersTable) {
      const rows = customersTable.querySelectorAll('tbody tr');
      for (const row of rows) {
        const cells = row.querySelectorAll('td');
        if (cells.length >= 1) {
          const name = cells[0].textContent.trim();
          if (name && name !== 'Заказчик' && name.length > 0 && !name.includes('Заказчик')) {
            let isMain = false;
            const checkbox = row.querySelector('.af-checkbox-v2.is-checked, input:checked, .is-checked');
            if (checkbox) {
              isMain = true;
            }
            
            if (customers.length === 0 && !isMain && rows.length > 0) {
              const anyChecked = Array.from(rows).some(r => r.querySelector('.af-checkbox-v2.is-checked, input:checked'));
              if (!anyChecked && customers.length === 0) {
                isMain = true;
              }
            }
            
            customers.push({ name: name, isMain: isMain });
          }
        }
      }
    }
    
    if (customers.length === 0) {
      const customerText = data['Заказчики'] || '';
      if (customerText && customerText !== '-') {
        const names = customerText.split(/[;,]/).map(s => s.trim()).filter(s => s && s !== '-');
        for (const name of names) {
          customers.push({ name: name, isMain: customers.length === 0 });
        }
      }
    }
    
    if (customers.length > 0) {
      data['Заказчики_массив'] = customers;
      data['Заказчики'] = customers.map(c => c.name).join('; ');
      const mainCustomer = customers.find(c => c.isMain);
      if (mainCustomer) {
        data['Главный_заказчик'] = mainCustomer.name;
      }
    }
    
    const okpd2Items = [];
    for (const table of allTables) {
      const header = table.querySelector('th');
      if (header && header.textContent.includes('Код ОКПД2')) {
        const rows = table.querySelectorAll('tbody tr');
        for (const row of rows) {
          const cells = row.querySelectorAll('td');
          if (cells.length >= 2) {
            const code = cells[0].textContent.trim();
            const name = cells[1].textContent.trim();
            if (code && name && code !== 'Код ОКПД2' && code.length > 0) {
              okpd2Items.push({ code: code, name: name });
            }
          }
        }
        break;
      }
    }
    
    if (okpd2Items.length > 0) {
      data['ОКПД2_массив'] = okpd2Items;
    }
    
    const okved2Items = [];
    for (const table of allTables) {
      const header = table.querySelector('th');
      if (header && header.textContent.includes('Код ОКВЭД2')) {
        const rows = table.querySelectorAll('tbody tr');
        for (const row of rows) {
          const cells = row.querySelectorAll('td');
          if (cells.length >= 2) {
            const code = cells[0].textContent.trim();
            const name = cells[1].textContent.trim();
            if (code && name && code !== 'Код ОКВЭД2' && code.length > 0) {
              okved2Items.push({ code: code, name: name });
            }
          }
        }
        break;
      }
    }
    
    if (okved2Items.length > 0) {
      data['ОКВЭД2_массив'] = okved2Items;
    }
    
  } catch (e) {
    console.error('Ошибка сбора:', e);
    data['Ошибка_сбора'] = e.message;
  }
  
  return data;
}

function extractCurrentTable() {
  try {
    const table = document.querySelector('table');
    if (!table) {
      return { headers: [], rows: [] };
    }
    
    const headers = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent.trim());
    const rows = Array.from(table.querySelectorAll('tbody tr')).map(tr => 
      Array.from(tr.querySelectorAll('td')).map(td => td.textContent.trim())
    );
    
    return { headers, rows };
  } catch (e) {
    return { headers: [], rows: [] };
  }
}

function getFileList() {
  const files = [];
  try {
    document.querySelectorAll('.files-list__item').forEach(item => {
      const nameEl = item.querySelector('.files-list__content .amphiris');
      if (nameEl) {
        const fileName = nameEl.textContent.trim();
        if (fileName && !files.includes(fileName)) {
          files.push(fileName);
        }
      }
    });
  } catch(e) {}
  return files;
}

function downloadFileByClick(fileName) {
  return new Promise((resolve) => {
    try {
      const items = document.querySelectorAll('.files-list__item');
      let targetItem = null;
      
      for (const item of items) {
        const nameEl = item.querySelector('.files-list__content .amphiris');
        if (nameEl && nameEl.textContent.trim() === fileName) {
          targetItem = item;
          break;
        }
      }
      
      if (!targetItem) {
        resolve(null);
        return;
      }
      
      const downloadBtn = targetItem.querySelector('.files-list__icon-wrapper, [class*="icon"]');
      if (!downloadBtn) {
        resolve(null);
        return;
      }
      
      const originalFetch = window.fetch;
      let downloadedData = null;
      
      window.fetch = function(...args) {
        const url = args[0];
        if (typeof url === 'string' && (url.includes('download') || url.includes('file') || url.includes('getFile'))) {
          return originalFetch.apply(this, args).then(async response => {
            const clone = response.clone();
            const blob = await clone.blob();
            const reader = new FileReader();
            reader.onloadend = () => {
              downloadedData = reader.result.split(',')[1];
            };
            reader.readAsDataURL(blob);
            return response;
          });
        }
        return originalFetch.apply(this, args);
      };
      
      downloadBtn.click();
      
      setTimeout(() => {
        window.fetch = originalFetch;
        resolve(downloadedData);
      }, 5000);
      
    } catch (e) {
      resolve(null);
    }
  });
}

document.getElementById('btn').addEventListener('click', async () => {
  const btn = document.getElementById('btn');
  const status = document.getElementById('status');
  const progress = document.getElementById('progress');

  btn.disabled = true;
  status.textContent = 'Обработка...';
  progress.style.display = 'block';

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab.url.includes('zakupki.rosatom.ru')) {
    status.textContent = 'Откройте сайт zakupki.rosatom.ru';
    status.className = 'err';
    btn.disabled = false;
    progress.style.display = 'none';
    return;
  }

  try {
    updateProgress(progress, 10, 'Переход к таблице...');
    
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: switchToTab,
      args: ['Детализированная потребность']
    });
    
    await delay(3000);
    
    updateProgress(progress, 25, 'Сбор таблицы...');
    
    const tableResult = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: extractCurrentTable,
    });
    
    const tableData = tableResult[0].result;
    
    updateProgress(progress, 40, 'Переход к основным данным...');
    
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: switchToTab,
      args: ['Основные данные']
    });
    
    await delay(3000);
    
    updateProgress(progress, 55, 'Сбор основных данных...');
    
    const mainDataResult = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: extractAllMainData,
    });
    
    const mainData = mainDataResult[0].result;
    
    updateProgress(progress, 70, 'Поиск файлов...');
    
    const filesResult = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: getFileList,
    });
    
    const fileNames = filesResult[0].result || [];
    
    updateProgress(progress, 75, 'Создание XML и Excel...');
    
    const xmlBlob = createXMLFile(mainData);
    const excelBlob = createExcelFile(tableData);
    
    const downloadedFiles = [];
    
    for (let i = 0; i < fileNames.length; i++) {
      updateProgress(progress, 75 + Math.floor((i / Math.max(fileNames.length, 1)) * 20), 
        `Файл ${i+1}/${fileNames.length}: ${fileNames[i].substring(0, 30)}`);
      
      const downloadResult = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: downloadFileByClick,
        args: [fileNames[i]]
      });
      
      if (downloadResult[0]?.result && downloadResult[0].result.length > 100) {
        downloadedFiles.push({
          name: fileNames[i],
          data: downloadResult[0].result
        });
      }
      
      await delay(500);
    }
    
    updateProgress(progress, 95, 'Создание ZIP архива...');
    
    const zip = new JSZip();
    zip.file(`Основные_данные.xml`, xmlBlob);
    
    if (excelBlob && tableData.rows && tableData.rows.length > 0) {
      zip.file(`Детализированная_потребность.xlsx`, excelBlob);
    }
    
    for (const file of downloadedFiles) {
      if (file.data && file.data.length > 0) {
        try {
          const binaryString = atob(file.data);
          const bytes = new Uint8Array(binaryString.length);
          for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
          }
          zip.file(file.name, bytes);
        } catch(e) {
          console.error('Ошибка добавления файла:', file.name, e);
        }
      }
    }
    
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const filename = `${sanitizeFilename(mainData['Предмет_закупки'] || 'Закупка')}.zip`;
    
    downloadBlob(zipBlob, filename);
    
    updateProgress(progress, 100, 'Готово');
    status.textContent = `Успешно. Позиций: ${tableData.rows?.length || 0}, Файлов: ${downloadedFiles.length}/${fileNames.length}, Поля в XML: ${Object.keys(mainData).length - 1}`;
    status.className = 'ok';
    
  } catch (e) {
    console.error('Ошибка:', e);
    status.textContent = 'Ошибка: ' + e.message;
    status.className = 'err';
  } finally {
    btn.disabled = false;
    setTimeout(() => { 
      if (progress) progress.style.display = 'none'; 
    }, 5000);
  }
});