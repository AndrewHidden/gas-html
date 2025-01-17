function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Wrap in HTML', 'wrapInHtmlTags')
      .addToUi();
}

function wrapInHtmlTags() {
  var body = DocumentApp.getActiveDocument().getBody();
  var numElements = body.getNumChildren();
  var htmlOutput = '';
  var imageCount = 1; // Счетчик изображений
  
  for (var i = 0; i < numElements; i++) {
    var element = body.getChild(i);
    
    // Обработка текстовых элементов
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var paragraph = element.asParagraph();
      var textElements = paragraph.getNumChildren();
      var paragraphHtml = '';

      for (var j = 0; j < textElements; j++) {
        var textElement = paragraph.getChild(j);
        
        // Проверка на текстовый элемент
        if (textElement.getType() === DocumentApp.ElementType.TEXT) {
          var text = textElement.asText();
          var content = text.getText();
          var formattedText = '';
          var currentStyle = '';
          var isBold = false;
          var isItalic = false;

          for (var k = 0; k < content.length; k++) {
            var char = content.charAt(k);
            var newBold = text.isBold(k);
            var newItalic = text.isItalic(k);
            var textColor = text.getForegroundColor();
            var backgroundColor = text.getBackgroundColor();

            // Формирование стиля для текущего символа
            var newStyle = '';
            if (textColor) {
              newStyle += 'color:' + textColor + ';';
            }
            if (backgroundColor) {
              newStyle += 'background-color:' + backgroundColor + ';';
            }

            // Если стиль изменился, закрываем предыдущие теги
            if (newBold !== isBold || newItalic !== isItalic || newStyle !== currentStyle) {
              // Закрыть предыдущие теги
              if (formattedText) {
                if (isBold) {
                  formattedText += '</strong>';
                }
                if (isItalic) {
                  formattedText += '</em>';
                }
                formattedText += '</span>';
              }

              // Открыть новые теги
              if (newBold) {
                formattedText += '<strong>';
              }
              if (newItalic) {
                formattedText += '<em>';
              }
              if (newStyle) {
                formattedText += '<span style="' + newStyle + '">';
              }

              // Обновить текущие значения стилей
              isBold = newBold;
              isItalic = newItalic;
              currentStyle = newStyle;
            }

            // Добавляем текущий символ
            formattedText += char;
          }

          // Закрываем оставшиеся теги
          if (formattedText) {
            if (isBold) {
              formattedText += '</strong>';
            }
            if (isItalic) {
              formattedText += '</em>';
            }
            if (currentStyle) {
              formattedText += '</span>';
            }
          }

          // Добавляем отформатированный текст в параграф
          paragraphHtml += formattedText;
        }
        // Проверка, если элемент - это изображение
        else if (textElement.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
          var imageName = 'image' + imageCount; // Создание уникального имени для изображения
          imageCount++; // Увеличиваем счетчик
            
          // Подготовка HTML для изображения по вашему формату
          paragraphHtml += `
          <p class="pimg">
            <picture>
              <source srcset="images/${imageName}.webp" type="image/webp">
              <img src="images/${imageName}.jpg" alt="${imageName}" width="" height="">
            </picture>
          </p>`;
        }
      }

      // Проверяем, не пустой ли параграф
      if (paragraphHtml.trim() !== '') {
        // Обернуть в <p> тег и добавить к общему HTML
        htmlOutput += `<p>${paragraphHtml}</p>`;
      }
    }

    // Обработка изображений, если они находятся отдельно от текстовых элементов
    else if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      var imageName = 'image' + imageCount; // Создание уникального имени для изображения
      imageCount++; // Увеличиваем счетчик

      // Подготовка HTML для изображения по вашему формату
      htmlOutput += `
      <p class="pimg">
        <picture>
          <source srcset="images/${imageName}.webp" type="image/webp">
          <img src="images/${imageName}.jpg" alt="${imageName}" width="600" height="377">
        </picture>
      </p>`;
    }
  }

  // Создание HTML-кода для боковой панели
  var html = `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; }
          textarea { width: 100%; height: 400px; }
        </style>
      </head>
      <body>
        <textarea readonly>${htmlOutput}</textarea>
      </body>
    </html>
  `;
  
  // Открытие боковой панели с результатом
  var output = HtmlService.createHtmlOutput(html)
      .setTitle('HTML Output')
      .setWidth(400);
  
  DocumentApp.getUi().showSidebar(output);
}