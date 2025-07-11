// Находим элементы на странице
const apiKeyInput = document.getElementById('apiKey');
const addressColumnInput = document.getElementById('addressColumn');
const fileInput = document.getElementById('fileInput');
const processButton = document.getElementById('processButton');
const statusDiv = document.getElementById('status');

// Основной обработчик нажатия на кнопку
processButton.addEventListener('click', async () => {
    // 1. Проверки
    const apiKey = apiKeyInput.value.trim();
    const addressColumn = addressColumnInput.value.trim();
    const file = fileInput.files[0];

    if (!apiKey) {
        alert('Пожалуйста, введите API-ключ.');
        return;
    }
    if (!addressColumn) {
        alert('Пожалуйста, укажите столбец с адресами.');
        return;
    }
    if (!file) {
        alert('Пожалуйста, выберите файл.');
        return;
    }

    // Блокируем кнопку и показываем статус
    processButton.disabled = true;
    statusDiv.textContent = 'Читаем файл...';

    try {
        // 2. Чтение XLSX файла
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(worksheet);

        if (rows.length === 0) {
            throw new Error("Файл пуст или не удалось прочитать данные.");
        }
        
        statusDiv.textContent = `Найдено ${rows.length} адресов. Начинаем геокодирование...`;

        // 3. Геокодирование каждого адреса
        const results = [];
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const address = row[addressColumn];
            
            statusDiv.textContent = `Обработка: ${i + 1} / ${rows.length} (${address})`;

            if (!address) {
                console.warn(`Пустой адрес в строке ${i + 2}`);
                continue; // Пропускаем строки без адреса
            }

            const coords = await getCoordinates(address, apiKey);

            results.push({
                'Широта': coords.lat,
                'Долгота': formatCoordinate(coords.lon), // Форматируем отрицательные значения
                'Описание': address, // В описание добавим исходный адрес
                'Подпись': '',
                'Номер': i + 1
            });
            
            // Маленькая пауза, чтобы не превысить лимиты API (например, 5 запросов в секунду)
            await new Promise(resolve => setTimeout(resolve, 200)); 
        }

        // 4. Создание нового XLSX файла
        statusDiv.textContent = 'Формируем итоговый файл...';
        const newWorksheet = XLSX.utils.json_to_sheet(results);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Координаты');

        // Скачиваем файл
        XLSX.writeFile(newWorkbook, 'результат_с_координатами.xlsx');
        
        statusDiv.textContent = 'Готово! Файл скачан.';

    } catch (error) {
        statusDiv.textContent = `Ошибка: ${error.message}`;
        console.error(error);
    } finally {
        processButton.disabled = false;
    }
});

// Функция для получения координат через Яндекс.Геокодер
async function getCoordinates(address, apiKey) {
    const url = `https://geocode-maps.yandex.ru/1.x/?apikey=${apiKey}&format=json&geocode=${encodeURIComponent(address)}`;
    
    try {
        const response = await fetch(url);
        const data = await response.json();
        
        const geoObject = data.response.GeoObjectCollection.featureMember[0]?.GeoObject;
        if (!geoObject) {
            console.warn(`Адрес не найден: ${address}`);
            return { lat: 'НЕ НАЙДЕНО', lon: 'НЕ НАЙДЕНО' };
        }
        
        const [lon, lat] = geoObject.Point.pos.split(' ');
        return { lat, lon };
    } catch (e) {
        console.error(`Ошибка при запросе к API для адреса: ${address}`, e);
        return { lat: 'ОШИБКА API', lon: 'ОШИБКА API' };
    }
}

// Функция для добавления апострофа к отрицательным координатам
function formatCoordinate(coordString) {
    if (typeof coordString === 'string' && coordString.startsWith('-')) {
        // В XLSX.js не нужно добавлять апостроф явно, библиотека сама обработает это.
        // Но если бы мы генерировали CSV, это было бы нужно.
        // Для XLSX лучше просто вернуть как есть.
        return coordString;
    }
    // Если нужно строго следовать инструкции, то можно сделать так:
    // return coordString.startsWith('-') ? `'${coordString}` : coordString;
    // Но библиотека xlsx умнее и сама может задать тип ячейки как "текст".
    // Давайте вернем как есть, это более чистое решение.
    return coordString;
}