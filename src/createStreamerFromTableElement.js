import CellStyle from './CellStyle.js'

export default function createStreamerFromTableElement(tableElement, skipEmptyRows = true) {

    const Q = selector => tableElement.querySelectorAll(selector)

    const streamRows = (elements, cellsSelector, callback, skipEmptyRows) => elements.forEach(rowElement => {
        const values = []
        const styles = []
        rowElement.querySelectorAll(cellsSelector).forEach(cellElement => {
            const isHidden = cellElement.dataset.xlHidden || false
            if (isHidden) {
                return
            }
            styles.push(new CellStyle(
                cellElement.dataset.xlHalign,
                cellElement.dataset.xlBold ? true : false
            ))
            values.push(cellElement.innerText) // type of value must be a string
        })
        if (!skipEmptyRows || values.length > 0) {
            callback(values, styles)
        }
    })

    return {

        async frozenPosition() {
            const rows = [ ...Q('caption'), ...Q('thead > tr') ]
            return { x: 0, y: rows.length }
        },

        async streamAll(callback) {
            streamRows([ tableElement ], 'caption', callback, true)
            streamRows(Q('thead > tr'), 'th', callback, skipEmptyRows)
            streamRows(Q('tbody > tr'), 'td', callback, skipEmptyRows)
            streamRows(Q('tfoot > tr'), 'td', callback, skipEmptyRows)
        }
    }
}
