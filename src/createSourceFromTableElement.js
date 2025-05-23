import CellStyle from './CellStyle.js'

export default function createSourceFromTableElement(tableElement, options = {}) {

    const { skipEmptyRows = true, author = '' } = options

    const Q = selector => tableElement.querySelectorAll(selector)

    const streamRows = (elements, cellsSelector, callback, skipEmptyRows, doComputeExtremes = true) => {
        for (const rowElement of elements) {
            const values = []
            const styles = []
            rowElement.querySelectorAll(cellsSelector).forEach(cellElement => {
                const opts = cellElement.dataset
                const isHidden = opts.xlHidden || false
                if (isHidden) {
                    return
                }
                const cellStyle = new CellStyle()
                const type = (/^(numeric|number|num|n)$/i.test(opts.xlType) ? CellStyle.TYPE_NUMERIC : CellStyle.TYPE_STRING)
                const formatCode = opts.xlFmt
                cellStyle.setType(type, formatCode)
                cellStyle.setFont(
                    opts.xlFont ? opts.xlFont : CellStyle.FONT_NAME_DEFAULT,
                    opts.xlFontSize ? opts.xlFontSize : CellStyle.FONT_SIZE_DEFAULT,
                    opts.xlColor,
                    opts.xlBold ? true : false,
                    opts.xlItalic ? true : false,
                    opts.xlUnderline ? true : false,
                    opts.xlStrikethrough ? true : false
                )
                cellStyle.setAlignment(
                    opts.xlHalign,
                    opts.xlValign,
                    opts.xlWrap
                )
                cellStyle.setBorderLeft(opts.xlBorderLeft || opts.xlBorder)
                cellStyle.setBorderRight(opts.xlBorderRight || opts.xlBorder)
                cellStyle.setBorderTop(opts.xlBorderTop || opts.xlBorder)
                cellStyle.setBorderBottom(opts.xlBorderBottom || opts.xlBorder)
                cellStyle.setBorderDiagonal(opts.xlBorderDiagonal)
                cellStyle.setFill(
                    opts.xlForegroundColor || opts.xlBackgroundColor
                        ? CellStyle.FILL_PATTERN_SOLID
                        : CellStyle.FILL_PATTERN_NONE,
                    opts.xlBackgroundColor
                )
                styles.push(cellStyle)
                values.push(cellElement.innerText) // type of value must be a string
            })
            if (!skipEmptyRows || values.length > 0) {
                if (callback({ values, styles, doComputeExtremes }))
                    return false
            }
        }
        return true
    }

    return {

        async getAuthor() {
            return author
        },

        async getFrozenPosition() {
            const rows = [ ...Q('caption'), ...Q('thead > tr') ]
            return { x: 0, y: rows.length }
        },

        async getReadableStream() {
            let canceled = false
            return new ReadableStream({
                start(controller) {
                    const callback = (rowData) => {
                        if (!canceled)
                            controller.enqueue(rowData)
                        return canceled
                    }
                    streamRows([ tableElement ], 'caption', callback, true, false)
                        && streamRows(Q('thead > tr'), 'th', callback, skipEmptyRows)
                        && streamRows(Q('tbody > tr'), 'td', callback, skipEmptyRows)
                        && streamRows(Q('tfoot > tr'), 'td', callback, skipEmptyRows)
                    controller.close()
                },
                pull(controller) { },
                cancel() {
                    canceled = true
                }
            })
        }
    }
}
