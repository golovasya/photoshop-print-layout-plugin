const { app, core } = require('photoshop');
const { storage, entrypoints } = require('uxp');
const fs = storage.localFileSystem;

// Глобальное состояние
let printsData = [];
let selectedPrintIndex = null;
let xlsxFilePath = null;

// Инициализация при загрузке панели
entrypoints.setup({
    panels: {
        printlayout: {
            create() {
                console.log('Print Layout Manager: Panel created');
                initializeUI();
            },
            show() {
                console.log('Print Layout Manager: Panel shown');
                setupLayerSelectionListener();
            },
            hide() {
                console.log('Print Layout Manager: Panel hidden');
            }
        }
    }
});

function initializeUI() {
    const loadXlsxBtn = document.getElementById('loadXlsxBtn');
    const runScriptBtn = document.getElementById('runScriptBtn');

    loadXlsxBtn.addEventListener('click', loadXlsxFile);
    runScriptBtn.addEventListener('click', runLayoutScript);

    showStatus('Готов к работе', 'success');
}

// Загрузка XLSX файла
async function loadXlsxFile() {
    try {
        showStatus('Выбор файла...', '');
        
        const file = await fs.getFileForOpening({
            types: ['xlsx', 'xls']
        });

        if (!file) {
            showStatus('Файл не выбран', 'error');
            return;
        }

        xlsxFilePath = file.nativePath;
        showStatus('Чтение файла...', '');

        // Чтение файла как ArrayBuffer
        const fileData = await file.read({ format: storage.formats.binary });
        
        // Парсинг XLSX с помощью SheetJS
        const workbook = XLSX.read(fileData, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        // Обработка данных (пропускаем заголовок)
        printsData = [];
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row[1] && !row[5]) continue; // Пропускаем пустые строки

            printsData.push({
                index: i - 1,
                photo: row[0] || '', // Столбец A - Фото
                size: row[1] || '', // Столбец B - Размер
                orderId: row[2] || '', // Столбец C - ID заказа
                name: row[3] || '', // Столбец D - Наименование
                color: row[4] || '', // Столбец E - Цвет
                article: row[5] || '', // Столбец F - Артикул
                realSize: extractRealSize(row[1]), // Реальный размер в мм
                layerId: null // Будет заполнено после раскладки
            });
        }

        displayFileInfo(file.name, printsData.length);
        renderPrintsList();
        showStatus(`Загружено принтов: ${printsData.length}`, 'success');

    } catch (error) {
        console.error('Error loading XLSX:', error);
        showStatus(`Ошибка: ${error.message}`, 'error');
    }
}

// Извлечение реального размера из строки размера
function extractRealSize(sizeStr) {
    if (!sizeStr) return '200x250';
    
    // Для детских размеров (122-152) и взрослых (XS-6XL)
    const childSizes = {
        '122': '190x220',
        '128': '200x230',
        '134': '210x240',
        '140': '220x250',
        '146': '230x260',
        '152': '240x270'
    };

    const adultSizes = {
        'XS': '200x250',
        'S': '210x260',
        'M': '220x270',
        'L': '230x280',
        'XL': '240x290',
        '2XL': '250x300',
        '3XL': '260x310',
        '4XL': '270x320',
        '5XL': '280x330',
        '6XL': '290x340'
    };

    // Извлекаем размер из строки типа "XS (40-42)" или "140"
    const match = sizeStr.match(/([XS0-9]+)/);
    if (match) {
        const size = match[1];
        return childSizes[size] || adultSizes[size] || '200x250';
    }

    return '200x250';
}

// Отображение информации о файле
function displayFileInfo(fileName, count) {
    const fileInfo = document.getElementById('fileInfo');
    fileInfo.style.display = 'block';
    fileInfo.textContent = `📄 ${fileName} — ${count} позиций`;
}

// Рендеринг списка принтов
function renderPrintsList() {
    const printsList = document.getElementById('printsList');
    printsList.innerHTML = '';

    printsData.forEach((print, index) => {
        const item = createPrintItem(print, index);
        printsList.appendChild(item);
    });
}

// Создание элемента принта
function createPrintItem(print, index) {
    const div = document.createElement('div');
    div.className = 'print-item';
    div.dataset.index = index;

    // Thumbnail
    const thumbnail = document.createElement('div');
    thumbnail.className = 'print-thumbnail';
    thumbnail.textContent = 'IMG';
    // TODO: загрузка реальных превью из таблицы если есть URL

    // Info
    const info = document.createElement('div');
    info.className = 'print-info';

    // Размер
    const sizeRow = document.createElement('div');
    sizeRow.className = 'print-info-row';
    sizeRow.innerHTML = `
        <span class="print-label">Размер:</span>
        <span class="print-value">${print.size}</span>
    `;

    // Артикул
    const articleRow = document.createElement('div');
    articleRow.className = 'print-info-row';
    articleRow.innerHTML = `
        <span class="print-label">Артикул:</span>
        <span class="print-value">${print.article}</span>
    `;

    // Реальный размер (редактируемый)
    const realSizeRow = document.createElement('div');
    realSizeRow.className = 'print-info-row';
    const sizeInput = document.createElement('input');
    sizeInput.type = 'text';
    sizeInput.className = 'size-input';
    sizeInput.value = print.realSize;
    sizeInput.addEventListener('change', (e) => {
        updatePrintSize(index, e.target.value);
    });

    realSizeRow.innerHTML = `<span class="print-label">Размер на листе:</span>`;
    realSizeRow.appendChild(sizeInput);

    info.appendChild(sizeRow);
    info.appendChild(articleRow);
    info.appendChild(realSizeRow);

    div.appendChild(thumbnail);
    div.appendChild(info);

    // Клик для выделения
    div.addEventListener('click', () => {
        selectPrintInUI(index);
        selectLayerInPhotoshop(print.layerId);
    });

    return div;
}

// Выделение принта в UI
function selectPrintInUI(index) {
    // Снимаем предыдущее выделение
    document.querySelectorAll('.print-item').forEach(item => {
        item.classList.remove('selected');
    });

    // Выделяем новый
    const item = document.querySelector(`[data-index="${index}"]`);
    if (item) {
        item.classList.add('selected');
        selectedPrintIndex = index;
    }
}

// Выделение слоя в Photoshop
async function selectLayerInPhotoshop(layerId) {
    if (!layerId) return;

    try {
        await core.executeAsModal(async () => {
            const doc = app.activeDocument;
            const layer = doc.layers.find(l => l.id === layerId);
            if (layer) {
                doc.activeLayers = [layer];
            }
        });
    } catch (error) {
        console.error('Error selecting layer:', error);
    }
}

// Слушатель выделения слоев в Photoshop
function setupLayerSelectionListener() {
    // TODO: Реализовать через события Photoshop API
    // В UXP пока нет прямых событий изменения выделения,
    // можно использовать периодическую проверку или notifier
}

// Обновление размера принта
async function updatePrintSize(index, newSize) {
    printsData[index].realSize = newSize;
    
    // Применяем новый размер к слою в Photoshop
    const layerId = printsData[index].layerId;
    if (!layerId) return;

    try {
        const [width, height] = newSize.split('x').map(s => parseFloat(s));
        if (!width || !height) {
            showStatus('Неверный формат размера (используйте ШИРИНАxВЫСОТА)', 'error');
            return;
        }

        await core.executeAsModal(async () => {
            const doc = app.activeDocument;
            const layer = doc.layers.find(l => l.id === layerId);
            
            if (layer) {
                // Конвертируем мм в пиксели (при 200 DPI)
                const dpi = doc.resolution;
                const widthPx = (width / 25.4) * dpi;
                const heightPx = (height / 25.4) * dpi;

                // Изменяем размер слоя
                const bounds = layer.bounds;
                const currentWidth = bounds.right - bounds.left;
                const currentHeight = bounds.bottom - bounds.top;

                const scaleX = (widthPx / currentWidth) * 100;
                const scaleY = (heightPx / currentHeight) * 100;

                layer.scale(scaleX, scaleY);
                
                showStatus(`Размер обновлен: ${newSize} мм`, 'success');
            }
        });
    } catch (error) {
        console.error('Error updating layer size:', error);
        showStatus(`Ошибка изменения размера: ${error.message}`, 'error');
    }
}

// Запуск скрипта раскладки
async function runLayoutScript() {
    try {
        showStatus('Запуск скрипта раскладки...', '');

        // Выбор файла скрипта
        const scriptFile = await fs.getFileForOpening({
            types: ['jsx', 'js']
        });

        if (!scriptFile) {
            showStatus('Скрипт не выбран', 'error');
            return;
        }

        // Чтение и выполнение скрипта
        const scriptContent = await scriptFile.read({ format: storage.formats.utf8 });
        
        await core.executeAsModal(async () => {
            // Выполнение ExtendScript в Photoshop
            await app.batchPlay([{
                _obj: 'AdobeScriptAutomation Scripts',
                javaScriptMessage: scriptContent
            }], {});
        });

        // После выполнения скрипта связываем слои с данными таблицы
        await linkLayersToData();

        showStatus('Скрипт выполнен успешно', 'success');

    } catch (error) {
        console.error('Error running script:', error);
        showStatus(`Ошибка выполнения скрипта: ${error.message}`, 'error');
    }
}

// Связывание слоев с данными из таблицы
async function linkLayersToData() {
    try {
        await core.executeAsModal(async () => {
            const doc = app.activeDocument;
            const layers = doc.layers;

            // Проходим по всем слоям и пытаемся связать с данными по артикулу
            layers.forEach(layer => {
                const layerName = layer.name;
                
                // Ищем соответствие по артикулу в имени слоя
                const matchedPrint = printsData.find(p => 
                    !p.layerId && layerName.includes(p.article)
                );

                if (matchedPrint) {
                    matchedPrint.layerId = layer.id;
                }
            });
        });

        renderPrintsList(); // Обновляем UI
    } catch (error) {
        console.error('Error linking layers:', error);
    }
}

// Показ статуса
function showStatus(message, type) {
    const status = document.getElementById('status');
    status.style.display = 'block';
    status.textContent = message;
    status.className = 'status';
    
    if (type === 'error') {
        status.classList.add('error');
    } else if (type === 'success') {
        status.classList.add('success');
    }

    // Автоскрытие через 5 секунд
    if (type) {
        setTimeout(() => {
            status.style.display = 'none';
        }, 5000);
    }
}
