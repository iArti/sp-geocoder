// !!! ВАЖНО: Замените эту строку на вашу часть ключа !!!
// Пример: если ваш ключ abcdef-1234-5678-90ab-cdefghijklMNO,
// то сюда нужно вставить "abcdef-1234-5678-90ab-cdefghijkl"
const API_KEY_PREFIX = "335c8ef1-a3dc-4c07-b1c7-64d4b431e";

// Находим элементы на странице
const apiKeySuffixInput = document.getElementById('apiKeySuffix'); // Обновленный ID
const addressColumnInput = document.getElementById('addressColumn');
const labelColumnInput = document.getElementById('labelColumn');
const fileInput = document.getElementById('fileInput');
const processButton = document.getElementById('processButton');
const statusDiv = document.getElementById('status');

// Основной обработчик нажатия на кнопку
processButton.addEventListener('click', async () => {
    // 1. Сборка ключа и проверки
    const apiKeySuffix = apiKeySuffixInput.value.trim();
    
    if (apiKeySuffix.length !== 3) {
        alert('Пожалуйста, введите ровно 3 последних символа вашего API-ключа.');
        return;
    }

    const fullApiKey = API_KEY_PREFIX + apiKeySuffix; // <-- Склеиваем ключ

    const addressColumn = addressColumnInput.value.trim();
    const labelColumn = labelColumnInput.value.trim();
    const file = fileInput.files[0];

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
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(worksheet);

        if (rows.length === 0) {
            throw new Error("Файл пуст или не удалось прочитать данные.");
        }
        
        statusDiv.textContent = `Найдено ${rows.length} адресов. Начинаем геокодирование...`;

        const results = [];
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const address = row[addressColumn];
            const label = labelColumn ? (row[labelColumn] || '') : '';
            
            statusDiv.textContent = `Обработка: ${i + 1} / ${rows.length} (${address})`;

            if (!address) {
                console.warn(`Пустой адрес в строке ${i + 2}`);
                results.push({
                    'Широта': 'АДРЕС НЕ УКАЗАН', 'Долгота': '', 'Описание': '', 'Подпись': label, 'Номер': i + 1
                });
                continue;
            }

            // Передаем в функцию уже полный ключ
            const coords = await getCoordinates(address, fullApiKey);

            results.push({
                'Широта': coords.lat,
                'Долгота': coords.lon,
                'Описание': address,
                'Подпись': label,
                'Номер': i + 1
            });
            
            await new Promise(resolve => setTimeout(resolve, 200)); 
        }

        statusDiv.textContent = 'Формируем итоговый файл...';
        const newWorksheet = XLSX.utils.json_to_sheet(results);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Координаты');
        XLSX.writeFile(newWorkbook, 'результат_с_координатами.xlsx');
        
        statusDiv.textContent = 'Готово! Файл скачан.';

    } catch (error) {
        statusDiv.textContent = `Ошибка: ${error.message}`;
        console.error(error);
    } finally {
        processButton.disabled = false;
    }
});

// Функция для получения координат. Она принимает уже полный ключ.
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