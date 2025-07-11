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
        const mode = document.querySelector('input[name="mode"]:checked').value;
        const apiKeySuffix = apiKeySuffixInput.value.trim();
        const addressColumnName = addressColumnInput.value.trim();
        const file = fileInput.files[0];

        if (apiKeySuffix.length !== 3) {
            alert('Пожалуйста, введите ровно 3 последних символа API-ключа Яндекса.');
            return;
        }
        const fullApiKey = API_KEY_PREFIX + apiKeySuffix; 

        if (!addressColumnName || !file) {
            alert('Пожалуйста, заполните все поля и выберите файл.');
            return;
        }
        
        processButton.disabled = true;
        statusDiv.textContent = `Начинаю обработку...`;
        
        const data = await file.arrayBuffer();
        
        // =========================================================================
        // НОВАЯ ЛОГИКА ОБРАБОТКИ
        // =========================================================================
        
        if (mode === 'google') {
            statusDiv.textContent = 'Режим Google: Модифицируем исходный файл...';
            console.log("РЕЖИМ GOOGLE: работаем напрямую с файлом для сохранения структуры.");

            // 1. Читаем файл, сохраняя всю структуру.
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            // 2. Находим букву нужного нам столбца (например, 'C').
            const columnLetter = findColumnLetter(worksheet, addressColumnName);
            if (!columnLetter) {
                throw new Error(`Столбец с названием "${addressColumnName}" не найден в файле.`);
            }
            console.log(`Найден столбец с адресами: ${columnLetter}`);

            // 3. Получаем диапазон строк, которые нужно обработать.
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            const startRow = range.s.r + 1; // Начинаем со второй строки (индекс 1)
            const endRow = range.e.r;

            // 4. Проходим по ячейкам ТОЛЬКО нужного столбца и заменяем их значения.
            for (let i = startRow; i <= endRow; i++) {
                const cellAddress = `${columnLetter}${i + 1}`;
                const cell = worksheet[cellAddress];
                
                // Если ячейка пуста или не содержит текста, пропускаем
                if (!cell || !cell.v) {
                    console.log(`Пропускаю пустую ячейку ${cellAddress}`);
                    continue;
                }
                
                const address = cell.v.toString();
                statusDiv.textContent = `Обработка строки ${i + 1}: ${address}`;
                console.log(`Обрабатываю ${cellAddress}: ${address}`);

                const coords = await getCoordinates(address, fullApiKey);
                const coordinateString = (coords.lat === 'НЕ НАЙДЕНО' || coords.lat === 'ОШИБКА API')
                    ? coords.lat
                    : `${coords.lat},${coords.lon}`;

                // Прямая модификация ячейки
                cell.v = coordinateString; // Заменяем значение
                cell.t = 's'; // Устанавливаем тип "строка", чтобы Excel не посчитал это формулой
                delete cell.w; // Удаляем старое отформатированное значение
                
                await new Promise(resolve => setTimeout(resolve, 200));
            }

            // 5. Сохраняем измененный workbook. Это сохранит всё форматирование.
            XLSX.writeFile(workbook, 'модифицированный_для_google.xlsx');

        } else { // Режим 'yandex' (старая, правильная для него логика)
            statusDiv.textContent = 'Режим Яндекс: Создаем новый файл...';
            const workbook = XLSX.read(data);
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
            const labelColumn = labelColumnInput.value.trim();
            const results = [];

            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const address = row[addressColumnName];
                const label = labelColumn ? (row[labelColumn] || '') : '';
                
                statusDiv.textContent = `Обработка строки ${i + 1}...`;

                let coords = { lat: 'НЕ НАЙДЕНО', lon: 'НЕ НАЙДЕНО' };
                if (address) {
                    coords = await getCoordinates(address, fullApiKey);
                }

                results.push({
                    'Широта': coords.lat, 'Долгота': coords.lon, 'Описание': address, 'Подпись': label, 'Номер': i + 1
                });
                await new Promise(resolve => setTimeout(resolve, 200));
            }

            const newWorksheet = XLSX.utils.json_to_sheet(results);
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Координаты');
            XLSX.writeFile(newWorkbook, 'для_яндекс_карт.xlsx');
        }

        statusDiv.textContent = 'Готово! Файл скачан.';
    } catch (error) {
        statusDiv.textContent = `Произошла ошибка! Детали в консоли (F12).`;
        console.error("КРИТИЧЕСКАЯ ОШИБКА:", error);
    } finally {
        processButton.disabled = false;
    }
});

// Вспомогательная функция для поиска буквы столбца по его названию
function findColumnLetter(worksheet, columnName) {
    const upperCaseColumnName = columnName.toUpperCase();
    
    // Проверяем, не ввел ли пользователь букву столбца напрямую (A, B, C...)
    if (/^[A-Z]+$/.test(upperCaseColumnName)) {
        return upperCaseColumnName;
    }

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: C });
        const cell = worksheet[cellAddress];
        if (cell && cell.v && cell.v.toString().trim().toUpperCase() === upperCaseColumnName) {
            return cellAddress.replace(/[0-9]/g, ''); // Из 'C1' получаем 'C'
        }
    }
    return null; // Столбец не найден
}

// Функция getCoordinates остается без изменений
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