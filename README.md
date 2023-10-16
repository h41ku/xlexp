xlexport
========

Простая библиотека для выгрузки табличных данных в Excel (xlsx) файл.

Возможности
-----------

- Использование как во frontend так и в backend
- Выгрузка данных в формате xlsx
- Закрепление указанной области
- Установка размера и стиля шрифта
- Установка стилей границ ячеек
- Установка цвета текста и фона ячеек
- Использование абстрактных асинхронных потоков данных
- Использование html-таблиц в качестве потоков данных

Пример использования
--------------------

```js
import * as xlexport from 'path/to/dist/xlexport.esm.min.js'

const tableElement = document.querySelector('table')
const buttonElement = document.querySelector('button.export')

buttonElement.addEventListener('click', async () => {
    buttonElement.setAttribute('disabled', true)
    const blob = await xlexport.exportToExcel(
        xlexport.createStreamerFromTableElement(tableElement)
    )
    xlexport.downloadAs(blob, "Пример выгрузки.xlsx")
    buttonElement.removeAttribute('disabled')
})
```
