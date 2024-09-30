const fileInput1 = document.getElementById('fileInput1');
const fileName1 = document.getElementById('fileName1');
const myForm = document.getElementById("myForm");
const okButton = document.getElementById("ok_btn");
const extButton = document.getElementById("exit_button");
const sendButton = document.getElementById("send");

fileInput1.addEventListener('change', function () {
    // Проверяем, есть ли выбранные файлы
    if (fileInput1.files.length > 0) {
        // Получаем имя первого выбранного файла
        const selectedFile = fileInput1.files[0].name;
        // Отображаем имя файла в span
        fileName1.textContent = selectedFile;
    }
});


function checkIn() {
    // Выключает кнопку до момента выбора обоих файлов
    const field1 = document.getElementById('fileInput1').value;
    const submitButton = document.getElementById('send');

    if (field1) {
        submitButton.disabled = false;
    }
    else {
        submitButton.disabled = true;
    }
}


function generateUUID() {
    // Генерирует UUID код, далее применяемый в назначении названия файлу готового отчета,
    // что позволяет избегать конфликтов при одновременной загрузке нескольких файлов
    var d = new Date().getTime();

    if (window.performance && typeof window.performance.now === "function") {
        d += performance.now();
    }

    var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var r = (d + Math.random() * 16) % 16 | 0;
        d = Math.floor(d / 16);
        return (c == 'x' ? r : (r & 0x3 | 0x8)).toString(16);
    });

    return uuid;
}

function handleClick() {
    // Очищает форму по нажатию на кнопку
    document.getElementById("popup-overlay").style.display = "none";
    fileName1.textContent = '';
    myForm.reset();
    sendButton.disabled = true;
    fileInput1.value = null;
    document.getElementById("nav").style.display = "flex"
}

function handleClickExit() {
    // Перезагружает страницу при нажатии кнопки 'Отмена'
    location.reload();
}

// Отправляет форму при нажатии на submit кнопку
myForm.addEventListener("submit", submitForm);

// Запускает функцию handleClick по нажатию на okButton
okButton.addEventListener("click", handleClick);

// Запускает функцию handleClick по нажатию extButton
extButton.addEventListener("click", handleClickExit);
