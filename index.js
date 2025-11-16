// =====================================================
// Print Layout Manager - UXP Plugin for Photoshop
// =====================================================

const { app } = require('photoshop');
const { storage, localFileSystem } = require('uxp').storage;
const fs = require('uxp').storage.localFileSystem;

// Глобальные переменные
let tableData = [];
let currentFile = null;
let selectedPrintIndex = null;
let layerToPrintMap = new Map(); // Соответствие слоёв к данным таблицы
let printToLayerMap = new Map(); // Обратное соответствие

// Элементы UI
let loadXlsxBtn, runScriptBtn, clearFileBtn;
let fileInfo, fileName, printsList, printDetails;
let searchInput, statusText, printCount;
let detailArticle, detailSize, detailColor, mockupImage;
let physicalWidth, physicalHeight, applySizeBtn;

// =====================================================
// Инициализация
// =====================================================

function init() {
    // Получаем элементы
    loadXlsxBtn = document.getElementById('loadXlsxBtn');
    runScriptBtn = document.getElementById('runScriptBtn');
    clearFileBtn = document.getElementById('clearFileBtn');
    fileInfo = document.getElementById('fileInfo');
    fileName = document.getElementById('fileName');
    printsList = document.getElementById('printsList');
    printDetails = document.getElementById('printDetails');
    searchInput = document.getElementById('searchInput');
    statusText = document.getElementById('statusText');
    printCount = document.getElementById('printCount');
    
    detailArticle = document.getElementById('detailArticle');
    detailSize = document.getElementById('detailSize');
    detailColor = document.getElementById('detailColor');
    mockupImage = document.getElementById('mockupImage');
    physicalWidth = document.getElementById('physicalWidth');
    physicalHeight = document.getElementById('physicalHeight');
    applySizeBtn = document.getElementById('applySizeBtn');

    // Привязываем обработчики
    loadXlsxBtn.addEventListener('click', loadXlsxFile);
    runScriptBtn.addEventListener('click', runLayoutScript);
    clearFileBtn.addEventListener('click', clearFile);
    searchInput.addEventListener('input', filterPrints);
    applySizeBtn.addEventListener('click', applyPhysicalSize);

    // Проверяем открытый документ
    checkDocument();
    
    // Обновляем список слоёв
    refreshPrintsList();
    
    updateStatus('Плагин готов к работе');
}

// =====================================================
// Загрузка XLSX файла
// =====================================================

async function loadXlsxFile() {
    try {
        updateStatus('Выбор файла...');
        
        const file = await fs.getFileForOpening({
            types: ['xlsx', 'xls']
        });

        if (!file) {
            updateStatus('Выбор файла отменён');
            return;
        }

        updateStatus('Чтение файла...');
        
        // Читаем файл как ArrayBuffer
        const arrayBuffer = await file.read({ format: storage.formats.binary });
        
        // Парсим XLSX с помощью SheetJS
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        // Берём первый лист
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Конвертируем в JSON
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        // Обрабатываем данные
        parseTableData(jsonData);
        
        currentFile = file;
        fileName.textContent = file.name;
        fileInfo.classList.remove('hidden');
        runScriptBtn.disabled = false;
        
        updateStatus(`Загружено ${tableData.length} записей из ${file.name}`);
        
        // Обновляем список
        refreshPrintsList();
        
    } catch (error) {
        console.error('Ошибка загрузки XLSX:', error);
        updateStatus('Ошибка: ' + error.message);
        showAlert('Ошибка загрузки файла', error.message);
    }
}

// =====================================================
// Парсинг данных таблицы
// =====================================================

function parseTableData(jsonData) {
    tableData = [];
    
    // Пропускаем заголовок (первая строка)
    for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        
        if (!row || row.length === 0) continue;
        
        const printData = {
            rowIndex: i,
            photo: row[0] || null,           // Колонка A (индекс 0) - Фото
            size: row[1] || 'Unknown',       // Колонка B (индекс 1) - Размер  
            orderId: row[2] || '',           // Колонка C (индекс 2) - ID заказа
            name: row[3] || '',              // Колонка D (индекс 3) - Наименование
            color: row[4] || '',             // Колонка E (индекс 4) - Цвет
            article: row[5] || 'Unknown',    // Колонка F (индекс 5) - Артикул продавца
            physicalWidth: null,
            physicalHeight: null,
            layerId: null                    // Будет заполнено при сопоставлении со слоями
        };
        
        tableData.push(printData);
    }
    
    console.log('Parsed table data:', tableData.length, 'records');
}

// =====================================================
// Очистка файла
// =====================================================

function clearFile() {
    currentFile = null;
    tableData = [];
    fileName.textContent = 'Файл не загружен';
    fileInfo.classList.add('hidden');
    runScriptBtn.disabled = true;
    refreshPrintsList();
    updateStatus('Файл очищен');
}

// =====================================================
// Запуск скрипта раскладки
// =====================================================

async function runLayoutScript() {
    try {
        updateStatus('Запуск скрипта раскладки...');
        
        // Здесь должна быть интеграция с твоим скриптом Сборщик v.3.5
        // Для демонстрации просто показываем сообщение
        
        await showAlert(
            'Запуск скрипта',
            'Интеграция со скриптом "Сборщик v.3.5" будет добавлена.\n\n' +
            'Скрипт должен:\n' +
            '1. Разместить принты на холсте\n' +
            '2. Присвоить слоям имена с артикулами\n' +
            '3. Вернуть управление плагину для синхронизации'
        );
        
        // После выполнения скрипта обновляем список
        refreshPrintsList();
        
        updateStatus('Скрипт выполнен');
        
    } catch (error) {
        console.error('Ошибка выполнения скрипта:', error);
        updateStatus('Ошибка: ' + error.message);
    }
}

// =====================================================
// Обновление списка принтов
// =====================================================

async function refreshPrintsList() {
    printsList.innerHTML = '';
    
    if (!app.activeDocument) {
        printsList.innerHTML = '<div class="hint" style="padding: 20px; text-align: center;">Нет открытого документа</div>';
        printCount.textContent = '0';
        return;
    }
    
    try {
        const doc = app.activeDocument;
        const layers = doc.layers;
        
        // Создаём соответствие между слоями и данными таблицы
        layerToPrintMap.clear();
        printToLayerMap.clear();
        
        let matchCount = 0;
        
        // Перебираем слои
        for (let i = 0; i < layers.length; i++) {
            const layer = layers[i];
            
            // Пропускаем фоновый слой
            if (layer.isBackgroundLayer) continue;
            
            // Ищем соответствие по артикулу в имени слоя
            const layerName = layer.name;
            
            for (let j = 0; j < tableData.length; j++) {
                const printData = tableData[j];
                
                // Проверяем, содержит ли имя слоя артикул
                if (layerName.includes(printData.article)) {
                    // Сохраняем ID слоя
                    printData.layerId = layer.id;
                    
                    // Получаем размеры слоя в мм
                    try {
                        const bounds = layer.bounds;
                        printData.physicalWidth = Math.round((bounds.right - bounds.left) * 0.352778 * 10) / 10; // px to mm
                        printData.physicalHeight = Math.round((bounds.bottom - bounds.top) * 0.352778 * 10) / 10;
                    } catch (err) {
                        console.error('Error getting layer bounds:', err);
                    }
                    
                    layerToPrintMap.set(layer.id, printData);
                    printToLayerMap.set(j, layer.id);
                    matchCount++;
                    break;
                }
            }
        }
        
        printCount.textContent = matchCount.toString();
        
        // Отображаем только сопоставленные принты
        const matchedPrints = tableData.filter(p => p.layerId !== null);
        
        if (matchedPrints.length === 0) {
            printsList.innerHTML = '<div class="hint" style="padding: 20px; text-align: center;">Нет сопоставленных слоёв.\nСлои должны содержать артикулы в названии.</div>';
            return;
        }
        
        matchedPrints.forEach((printData, index) => {
            const item = createPrintItem(printData, index);
            printsList.appendChild(item);
        });
        
        updateStatus(`Найдено ${matchCount} принтов на холсте`);
        
    } catch (error) {
        console.error('Error refreshing prints list:', error);
        printsList.innerHTML = '<div class="hint" style="padding: 20px; text-align: center; color: red;">Ошибка: ' + error.message + '</div>';
    }
}

// =====================================================
// Создание элемента принта
// =====================================================

function createPrintItem(printData, index) {
    const item = document.createElement('div');
    item.className = 'print-item';
    item.dataset.index = index;
    item.dataset.layerId = printData.layerId;
    
    // Миниатюра (пока заглушка)
    const thumbnail = document.createElement('div');
    thumbnail.className = 'print-thumbnail';
    thumbnail.innerHTML = '<span style="font-size: 20px;">🖼️</span>';
    
    // Информация
    const info = document.createElement('div');
    info.className = 'print-info';
    
    const article = document.createElement('div');
    article.className = 'print-article';
    article.textContent = printData.article;
    
    const meta = document.createElement('div');
    meta.className = 'print-meta';
    
    const sizeBadge = document.createElement('span');
    sizeBadge.className = 'print-size-badge';
    sizeBadge.textContent = printData.size;
    
    const dimensions = document.createElement('span');
    if (printData.physicalWidth && printData.physicalHeight) {
        dimensions.textContent = `${printData.physicalWidth}×${printData.physicalHeight} мм`;
    } else {
        dimensions.textContent = 'Размер не определён';
    }
    
    meta.appendChild(sizeBadge);
    meta.appendChild(dimensions);
    
    info.appendChild(article);
    info.appendChild(meta);
    
    item.appendChild(thumbnail);
    item.appendChild(info);
    
    // Обработчик клика
    item.addEventListener('click', () => selectPrint(index, printData));
    
    return item;
}

// =====================================================
// Выбор принта
// =====================================================

async function selectPrint(index, printData) {
    selectedPrintIndex = index;
    
    // Обновляем UI
    document.querySelectorAll('.print-item').forEach(item => {
        item.classList.remove('selected');
    });
    
    const selectedItem = document.querySelector(`[data-index="${index}"]`);
    if (selectedItem) {
        selectedItem.classList.add('selected');
    }
    
    // Показываем детали
    showPrintDetails(printData);
    
    // Выделяем слой в Photoshop
    try {
        if (printData.layerId && app.activeDocument) {
            const layer = app.activeDocument.layers.find(l => l.id === printData.layerId);
            if (layer) {
                app.activeDocument.activeLayers = [layer];
                updateStatus(`Выбран: ${printData.article}`);
            }
        }
    } catch (error) {
        console.error('Error selecting layer:', error);
    }
}

// =====================================================
// Показ деталей принта
// =====================================================

function showPrintDetails(printData) {
    printDetails.classList.remove('hidden');
    
    detailArticle.textContent = printData.article;
    detailSize.textContent = printData.size;
    detailColor.textContent = printData.color || 'Не указан';
    
    physicalWidth.value = printData.physicalWidth || '';
    physicalHeight.value = printData.physicalHeight || '';
    
    // Мокап - пока заглушка
    mockupImage.src = '';
    mockupImage.alt = 'Мокап недоступен';
}

// =====================================================
// Применение физического размера
// =====================================================

async function applyPhysicalSize() {
    if (selectedPrintIndex === null) {
        await showAlert('Ошибка', 'Сначала выберите принт из списка');
        return;
    }
    
    const width = parseFloat(physicalWidth.value);
    const height = parseFloat(physicalHeight.value);
    
    if (isNaN(width) || isNaN(height) || width <= 0 || height <= 0) {
        await showAlert('Ошибка', 'Введите корректные размеры (мм)');
        return;
    }
    
    try {
        const printData = tableData.find(p => p.layerId !== null)[selectedPrintIndex];
        
        if (!printData || !printData.layerId) {
            await showAlert('Ошибка', 'Слой не найден');
            return;
        }
        
        const doc = app.activeDocument;
        const layer = doc.layers.find(l => l.id === printData.layerId);
        
        if (!layer) {
            await showAlert('Ошибка', 'Слой не найден в документе');
            return;
        }
        
        // Конвертируем мм в пиксели (72 DPI)
        const widthPx = width / 0.352778;
        const heightPx = height / 0.352778;
        
        // Получаем текущие размеры
        const bounds = layer.bounds;
        const currentWidth = bounds.right - bounds.left;
        const currentHeight = bounds.bottom - bounds.top;
        
        // Вычисляем масштаб
        const scaleX = (widthPx / currentWidth) * 100;
        const scaleY = (heightPx / currentHeight) * 100;
        
        // Применяем масштабирование
        await layer.scale(scaleX, scaleY);
        
        // Обновляем данные
        printData.physicalWidth = width;
        printData.physicalHeight = height;
        
        updateStatus(`Размер изменён: ${width}×${height} мм`);
        
        // Обновляем список
        refreshPrintsList();
        
    } catch (error) {
        console.error('Error applying size:', error);
        await showAlert('Ошибка', 'Не удалось применить размер: ' + error.message);
    }
}

// =====================================================
// Фильтрация принтов
// =====================================================

function filterPrints() {
    const query = searchInput.value.toLowerCase();
    
    document.querySelectorAll('.print-item').forEach(item => {
        const article = item.querySelector('.print-article').textContent.toLowerCase();
        
        if (article.includes(query)) {
            item.style.display = 'flex';
        } else {
            item.style.display = 'none';
        }
    });
}

// =====================================================
// Проверка документа
// =====================================================

function checkDocument() {
    if (!app.activeDocument) {
        updateStatus('Нет открытого документа');
    }
}

// =====================================================
// Утилиты
// =====================================================

function updateStatus(message) {
    statusText.textContent = message;
    console.log('Status:', message);
}

async function showAlert(title, message) {
    const { app: uxpApp } = require('photoshop');
    const options = {
        title: title,
        message: message
    };
    
    try {
        await uxpApp.showAlert(message);
    } catch (e) {
        console.log(title + ': ' + message);
    }
}

// =====================================================
// Запуск при загрузке
// =====================================================

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}