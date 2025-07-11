console.log("✅ Скрипт script.js успешно загружен и готов к работе.");

// !!! ВАЖНО: Убедитесь, что здесь ваша реальная часть ключа !!!
const API_KEY_PREFIX = "335c8ef1-a3dc-4c07-b1c7-64d4b431e";

// Находим элементы на странице
const apiKeySuffixInput = document.getElementById('apiKeySuffix');
const addressColumnInput = document.getElementById('addressColumn');
const labelColumnInput = document.getElementById('labelColumn');
const fileInput = document.getElementById('fileInput');
const processButton = document.getElementById('processButton');
const statusDiv = document.getElementById('status');

// Основной обработчик
processButton.addEventListener('click', async () => {
    try {
        // 1. Сбор данных из формы
        const mode = document.querySelector('input[name="mode"]:checked').value;
        const apiKeySuffix = apiKeySuffixInput.value.trim();
        const addressColumn = addressColumnInput.value.trim();
        const labelColumn = labelColumnInput.value.trim();
        const file = fileInput.files[0];

        if (apiKeySuffix.length !== 3) {
            alert('Пожалуйста, введите ровно 3 последних символа API-ключа Яндекса.');
            return;
        }
        
        // Ключ Яндекса собирается и используется в ЛЮБОМ режиме.
        const fullApiKey = API_KEY_PREFIX + apiKeySuffix; 

        if (!addressColumn || !file) {
            alert('Пожалуйста, заполните все поля и выберите файл.');
            return;
        }
        
        processButton.disabled = true;
        statusDiv.textContent = `Начинаю обработку для ${mode === 'google' ? 'Google Карт' : 'Яндекс.Карт'}...`;
        
        // 2. Чтение файла
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        if (rows.length === 0) throw new Error("Файл пуст.");

        // 3. Геокодирование и формирование результата
        const results = [];
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const address = row[addressColumn];
            const label = labelColumn ? (row[labelColumn] || '') : '';
            
            statusDiv.textContent = `Обработка: ${i + 1} / ${rows.length}...`;

            let coords = { lat: 'НЕ НАЙДЕНО', lon: 'НЕ НАЙДЕНО' };
            if (address) {
                // ВЫЗОВ ГЕОКОДЕРА ЯНДЕКСА ПРОИСХОДИТ ЗДЕСЬ, НЕЗАВИСИМО ОТ РЕЖИМА
                coords = await getCoordinates(address, fullApiKey);
            }

            // А ЗДЕСЬ МЫ ТОЛЬКО ФОРМАТИРУЕМ УЖЕ ПОЛУЧЕННЫЙ РЕЗУЛЬТАТ
            if (mode === 'google') {
                results.push({
                    'Latitude': coords.lat,       // Широта с понятным для Google названием
                    'Longitude': coords.lon,      // Долгота с понятным для Google названием
                    'Name': label || address,     // Название метки
                    'Description': address        // Описание метки
                });
            } else { // Режим 'yandex'
                results.push({
                    'Широта': coords.lat,
                    'Долгота': coords.lon,
                    'Описание': address,
                    'Подпись': label,
                    'Номер': i + 1
                });
            }
            
            await new Promise(resolve => setTimeout(resolve, 200));
        }

        // 4. Создание и скачивание файла
        const newWorksheet = XLSX.utils.json_to_sheet(results);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Координаты');
        const fileName = mode === 'google' ? 'для_google_maps.xlsx' : 'для_яндекс_карт.xlsx';
        XLSX.writeFile(newWorkbook, fileName);
        
        statusDiv.textContent = 'Готово! Файл скачан.';
    } catch (error) {
        statusDiv.textContent = `Произошла ошибка! Детали в консоли (F12).`;
        console.error("КРИТИЧЕСКАЯ ОШИБКА:", error);
    } finally {
        processButton.disabled = false;
    }
});

// Эта функция — "мозг". Она всегда обращается к Яндексу.
async function getCoordinates(address, apiKey) {
    const url = `https://geocode-maps.yandex.ru/1.x/?apikey=${apiKey}&format=json&geocode=${encodeURIComponent(address)}`;
    try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`Сетевая ошибка: ${response.status}`);
        const data = await response.json();
        if (data.error) throw new Error(`Ошибка API Яндекса: ${data.error.message}`);
        
        const geoObject = data.response?.GeoObjectCollection.featureMember[0]?.GeoObject;
        if (!geoObject) {
            console.warn(`Адрес не найден на карте: ${address}`);
            return { lat: 'НЕ НАЙДЕНО', lon: 'НЕ НАЙДЕНО' };
        }
        
        const [lon, lat] = geoObject.Point.pos.split(' ');
        return { lat, lon };
    } catch (e) {
        console.error(`Ошибка при запросе к API для адреса: ${address}`, e);
        return { lat: 'ОШИБКА API', lon: 'ОШИБКА API' };
    }
}