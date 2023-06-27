import CellStyle from './CellStyle.js'

export default function createStreamerFromTableElement(tableElement) {

    const trElements = tableElement.querySelectorAll('tr')

    return {

        async streamAll(callback) {
            trElements.forEach(trElement => {
                const values = []
                const styles = []
                trElement.querySelectorAll('th, td').forEach(tdElement => {
                    const isHidden = tdElement.dataset.xlHidden || false
                    if (isHidden) {
                        return
                    }
                    styles.push(new CellStyle(
                        tdElement.dataset.xlHalign,
                        tdElement.dataset.xlBold ? true : false
                    ))
                    values.push(tdElement.innerText) // type of value must be a string
                })
                callback(values, styles)
            })
        }
    }
}
