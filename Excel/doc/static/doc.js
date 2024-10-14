function copyToClipboard(button) {
    // Находит <pre> в том же <div>, что и кнопка
    var preElement = button.parentNode.querySelector('pre');
    
    // Создает временное текстовое поле
    var textarea = document.createElement('textarea');
    textarea.value = preElement.textContent;
    
    // Добавляет это текстовое поле в <body>
    document.body.appendChild(textarea);
    
    // Выделяет текст
    textarea.select();

    // Копирование через execCommand
    try {
        document.execCommand('copy');
    } catch (err) {
        alert('Не удалось скоприровать элемент');
    }

    // Удаляет созданное на время текстовое поле
    document.body.removeChild(textarea);
}