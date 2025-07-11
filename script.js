// =================== ОТЛАДКА ===================
// Это сообщение появится в консоли (F12), как только страница загрузится.
// Если вы его не видите, значит, файл script.js не подключился к index.html
console.log("✅ Скрипт script.js успешно загружен и готов к работе.");
// ===============================================


// !!! ВАЖНО: Убедитесь, что здесь ваша реальная часть ключа !!!
const API_KEY_PREFIX = "335c8ef1-a3dc-4c07-b1c7-64d4b431e";


// Находим элементы на странице
const apiKeySuffixInput = document.getElementById('apiKeySuffix');
const addressColumnInput = document.getElementById('addressColumn');
const labelColumnInput = document.getElementById('labelColumn');
const fileInput = document.getElementById('fileInput');
const processButton = document.getElementById('processButton');
const statusDiv = document.getElementById('status');


// Основной обработчик нажатия на кнопку
processButton.addEventListener('click', async () => {
    // try...catch - это ловушка для ошибок. Если что-то пойдет не так внутри 'try',
    // выполнение перепрыгнет в блок 'catch', и мы увидим ошибку в консоли.
    try {
        console.log("--- [СТАРТ] --- Кнопка 'Обработать' нажата.");

        // 1. Сборка ключа и проверки
        const apiKeySuffix = apiKeySuffixInput.value.trim();
        const addressColumn = addressColumnInput.value.trim();
        const labelColumn = labelColumnInput.value.trim();
        const file = fileInput.files[0];
        
        console.log("Шаг 1: Получены данные из формы:", {
            apiKeySuffix,
            addressColumn,
            labelColumn,
            fileExists: !!file
        });

        if (apiKeySuffix.length !== 3) {
            alert('Пожалуйста, введите ровно 3 последних символа вашего API-ключа.');
            console.error("ОШИБКА ВВОДА: Неверная длина суффикса ключа.");
            return; // Прекращаем выполнение
        }
        
        const fullApiKey = API_KEY_PREFIX + apiKeySuffix;
        if (API_KEY_PREFIX === "ВАША_ОСНОВНАЯ_ЧАСТЬ_КЛЮЧА_БЕЗ_ПОСЛЕДНИХ_3_СИМВОЛОВ") {
             alert('ВНИМАНИЕ: Вы не заменили базовую часть API-ключа в файле script.js!');
             console.warn("ПРЕДУПРЕЖДЕНИЕ: Используется ключ-заглушка.");
        }
        console.log("Ключ API собран.");

        if (!addressColumn) {
            alert('Пожалуйста, укажите столбец с адресами.');
            console.error("ОШИБКА ВВОДА: Не указан столбец с адресами.");
            return;
        }
        if (!file) {
            alert('Пожалуйста, выберите файл.');
            console.error("ОШИБКА ВВОДА: Файл не выбран.");
            return;
        }

        processButton.disabled = true;
        statusDiv.textContent = 'Начинаю обработку... (подробности в консоли F12)';

        // 2. Чтение XLSX файла
        console.log("Шаг 2: Чтение файла...");
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(worksheet);
        console.log(`Файл успешно прочитан. Найдено строк: ${rows.length}`);

        if (rows.length === 0) throw new Error("Файл пуст или имеет неверный формат.");
        
        // 3. Геокодирование
        console.log("Шаг 3: Начало цикла геокодирования...");
        const results = [];
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const address = row[addressColumn];
            const label = labelColumn ? (row[labelColumn] || '') : '';
            
            statusDiv.textContent = `Обработка: ${i + 1} / ${rows.length}...`;
            console.log(`-> Обрабатываю строку ${i + 1}: Адрес='${address}', Подпись='${label}'`);

            if (!address) {
                console.warn(`   Пропускаю строку ${i + 1}, т.к. адрес пуст.`);
                results.push({'Широта': 'АДРЕС НЕ УКАЗАН', 'Долгота': '', 'Описание': '', 'Подпись': label, 'Номер': i + 1});
                continue;
            }

            const coords = await getCoordinates(address, fullApiKey);
            console.log(`   <- Получен результат для '${address}':`, coords);

            results.push({'Широта': coords.lat, 'Долгота': coords.lon, 'Описание': address, 'Подпись': label, 'Номер': i + 1});
            await new Promise(resolve => setTimeout(resolve, 200)); // Пауза
        }

        // 4. Создание нового файла
        console.log("Шаг 4: Формирование итогового XLSX файла...");
        const newWorksheet = XLSX.utils.json_to_sheet(results);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Координаты');
        XLSX.writeFile(newWorkbook, 'результат_с_координатами.xlsx');
        
        statusDiv.textContent = 'Готово! Файл скачан.';
        console.log("--- [УСПЕХ] --- Обработка завершена.");

    } catch (error) {
        // Если на любом из этапов выше произойдет ошибка, мы попадем сюда
        statusDiv.textContent = `Произошла ошибка! Смотрите детали в консоли (F12).`;
        console.error("--- [КРИТИЧЕСКАЯ ОШИБКА] --- Выполнение прервано:", error);
    } finally {
        // Этот блок выполнится всегда: и после успеха, и после ошибки
        processButton.disabled = false;
        console.log("--- [КОНЕЦ] --- Кнопка снова активна.");
    }
});


// Функция для получения координат. Она остается без изменений.
async function getCoordinates(address, apiKey) {
    const url = `https://geocode-maps.yandex.ru/1.x/?apikey=${apiKey}&format=json&geocode=${encodeURIComponent(address)}`;
    try {
        const response = await fetch(url);
        if (!response.ok) { // Проверка статуса ответа от сервера
            throw new Error(`Сетевая ошибка: ${response.status} ${response.statusText}`);
        }
        const data = await response.json();
        
        // Добавим проверку на ошибку в ответе от Яндекса (например, невалидный ключ)
        if (data.error) {
            throw new Error(`Ошибка API Яндекса: ${data.error.message}`);
        }
        
        const geoObject = data.response?.GeoObjectCollection.featureMember[0]?.GeoObject;
        if (!geoObject) {
            console.warn(`   Адрес не найден на карте: ${address}`);
            return { lat: 'НЕ НАЙДЕНО', lon: 'НЕ НАЙДЕНО' };
        }
        
        const [lon, lat] = geoObject.Point.pos.split(' ');
        return { lat, lon };
    } catch (e) {
        console.error(`   Ошибка при запросе к API для адреса: ${address}`, e);
        return { lat: 'ОШИБКА API', lon: 'ОШИБКА API' };
    }
}