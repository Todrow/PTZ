const fileInput1 = document.getElementById('fileInput1');
const fileName1 = document.getElementById('fileName1');
const fileInput2 = document.getElementById('fileInput2');
const fileName2 = document.getElementById('fileName2');
const myForm = document.getElementById("myForm");
const okButton = document.getElementById("ok_btn");
const extButton = document.getElementById("exit_button");

fileInput1.addEventListener('change', function () {
    // Проверяем, есть ли выбранные файлы
    if (fileInput1.files.length > 0) {
        // Получаем имя первого выбранного файла
        const selectedFile = fileInput1.files[0].name;
        // Отображаем имя файла в span
        fileName1.textContent = selectedFile;
    } else {
        // Если файл не выбран, изменяем текст на стандартный
        fileName1.textContent = 'Выберите файл';
    }
});

fileInput2.addEventListener('change', function () {
    // Проверяем, есть ли выбранные файлы
    if (fileInput2.files.length > 0) {
        // Получаем имя первого выбранного файла
        const selectedFile = fileInput2.files[0].name;
        // Отображаем имя файла в span
        fileName2.textContent = selectedFile;
    } else {
        // Если файл не выбран, изменяем текст на стандартный
        fileName2.textContent = 'Выберите файл';
    }
});

function generateUUID() {
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
    document.getElementById("popup-overlay").style.display = "none";
}

function handleClickExit() {
    location.reload();
}

myForm.addEventListener("submit", submitForm);

okButton.addEventListener("click", handleClick);

extButton.addEventListener("click", handleClickExit);