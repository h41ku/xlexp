# xlexp

Простая библиотека для выгрузки табличных данных в Excel (xlsx) файл.

## Возможности

- Использование как во frontend так и в backend
- Выгрузка данных в формате xlsx
- Закрепление указанной области
- Возможность указать автора
- Установка размера и стиля шрифта
- Установка стилей границ ячеек
- Установка цвета текста и фона ячеек
- Использование `Streams API` для работы с асинхронными потоками данных
- Использование HTML-таблиц в качестве потоков данных

## Пример использования

Выгрузка таблицы:

```html
<table class="sample">
    <caption>Население некоторых городов</caption>
    <thead>
        <tr>
            <th data-xl-bold="on" data-xl-border="thin">Идентификатор</th>
            <th data-xl-bold="on" data-xl-border="thin">Город</th>
            <th data-xl-bold="on" data-xl-border="thin" data-xl-halign="right">Население</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>78123</td>
            <td>Урюпинск</td>
            <td data-xl-halign="right">23709</td>
        </tr>
        <tr>
            <td>34224</td>
            <td>Тюмень</td>
            <td data-xl-halign="right">992311</td>
        </tr>
        <tr>
            <td>22333</td>
            <td>Ульяновск</td>
            <td data-xl-halign="right">2349833</td>
        </tr>
    </tbody>
    <tfoot>
        <tr>
            <td></td>
            <td>Среднее</td>
            <td data-xl-halign="right">1121951</td>
        </tr>
    </tfoot>
</table>
```

Создайте источник, используя DOM-элемент, и выгрузите данные:


```js
import { exportToExcel, createSourceFromTableElement } from 'xlexp'
import passIt from 'pass-it'

const tableElement = document.querySelector('table')
const buttonElement = document.querySelector('button.export')

buttonElement.addEventListener('click', async () => {
    buttonElement.setAttribute('disabled', true)
    const blob = await exportToExcel(
        createSourceFromTableElement(tableElement)
    )
    passIt(blob, { download: "Пример выгрузки.xlsx" })
    buttonElement.removeAttribute('disabled')
})
```

Вы можете создавать источники с помощью своей функции, возвращающей объект,
соответствующий интерфейсу `WorksheetSource`:

```ts
export interface WorksheetSource {
    getAuthor(): Promise<string>;
    getFrozenPosition(): Promise<Position>;
    getReadableStream(): Promise<ReadableStream<Row>>;
};
```
