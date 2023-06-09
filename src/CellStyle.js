const HORIZONTAL_ALIGNMENT_LEFT = 'left'
const HORIZONTAL_ALIGNMENT_RIGHT = 'right'

class CellStyle {

    constructor(horizontalAlignment = HORIZONTAL_ALIGNMENT_LEFT, isBold = false) {
        this.horizontalAlignment = horizontalAlignment
        this.isBold = isBold
    }

    isEquals(cellStyle) {
        return this.horizontalAlignment === cellStyle.horizontalAlignment
            && this.isBold === cellStyle.isBold
    }
}

CellStyle.HORIZONTAL_ALIGNMENT_LEFT = HORIZONTAL_ALIGNMENT_LEFT
CellStyle.HORIZONTAL_ALIGNMENT_RIGHT = HORIZONTAL_ALIGNMENT_RIGHT

export default CellStyle
