const fileInput1 = document.getElementById('fileInput1');
const fileName1 = document.getElementById('fileName1');
const fileInput2 = document.getElementById('fileInput2');
const fileName2 = document.getElementById('fileName2');
const myForm = document.getElementById("myForm");

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


myForm.addEventListener("submit", function(event) {
    event.preventDefault()
    document.getElementById("popup-overlay").style.display = "block";
    myForm.reset();
});


function handleClick() {
    document.getElementById("popup-overlay").style.display = "none";
}

document.getElementById("ok_btn").addEventListener("click", handleClick);
