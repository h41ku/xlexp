const FONT_NAME_DEFAULT = 'Calibri'

const FONT_SIZE_DEFAULT = 11

const HORIZONTAL_ALIGNMENT_DEFAULT = null
const HORIZONTAL_ALIGNMENT_LEFT = 'left'
const HORIZONTAL_ALIGNMENT_CENTER = 'center'
const HORIZONTAL_ALIGNMENT_RIGHT = 'right'

const VERTICAL_ALIGNMENT_DEFAULT = null
const VERTICAL_ALIGNMENT_TOP = 'top'
const VERTICAL_ALIGNMENT_CENTER = 'center'
const VERTICAL_ALIGNMENT_BOTTOM = 'bottom'

const BORDER_THICKNESS_NONE = null
const BORDER_THICKNESS_THIN = 'thin'

const FILL_PATTERN_NONE = 'none'
const FILL_PATTERN_SOLID = 'solid'

const TYPE_STRING = 's'
const TYPE_NUMERIC = 'n'

const FORMAT_CODE_GENERAL = 'General'

const COLOR_DEFAULT = null

const fontsAreEquals = (a, b) => (
    a.name === b.name
        && a.size === b.size
        && a.color === b.color
        && a.bold === b.bold
        && a.italic === b.italic
        && a.underline === b.underline
        && a.strikethrough === b.strikethrough
)

const alignmentsAreEquals = (a, b) => (
    a.horizontal === b.horizontal
        && a.vertical === b.vertical
        && a.wrapText === b.wrapText
)

const bordersAreEquals = (a, b) => (
    a.thickness === b.thickness
        && a.color === b.color
)

const fillsAreEquals = (a, b) => (
    a.pattern === b.pattern
        && a.bgColor === b.bgColor
)

class CellStyle {

    constructor() {
        this.setFont()
        this.setAlignment()
        this.setBorderLeft()
        this.setBorderRight()
        this.setBorderTop()
        this.setBorderBottom()
        this.setBorderDiagonal()
        this.setFill()
        this.setType()
    }

    setFont(name = FONT_NAME_DEFAULT, size = FONT_SIZE_DEFAULT, color = COLOR_DEFAULT, bold = false, italic = false, underline = false, strikethrough = false) {
        this.font = { name, size, color, bold, italic, underline, strikethrough }
    }

    setAlignment(horizontal = HORIZONTAL_ALIGNMENT_DEFAULT, vertical = VERTICAL_ALIGNMENT_DEFAULT, wrapText = false) {
        this.alignment = { horizontal, vertical, wrapText }
    }

    setBorderLeft(thickness = BORDER_THICKNESS_NONE, color = 64) {
        this.borderLeft = { thickness, color }
    }

    setBorderRight(thickness = BORDER_THICKNESS_NONE, color = 64) {
        this.borderRight = { thickness, color }
    }

    setBorderTop(thickness = BORDER_THICKNESS_NONE, color = 64) {
        this.borderTop = { thickness, color }
    }

    setBorderBottom(thickness = BORDER_THICKNESS_NONE, color = 64) {
        this.borderBottom = { thickness, color }
    }

    setBorderDiagonal(thickness = BORDER_THICKNESS_NONE, color = 64) {
        this.borderDiagonal = { thickness, color }
    }

    setFill(pattern = FILL_PATTERN_NONE, bgColor = COLOR_DEFAULT) {
        this.fill = { pattern, bgColor }
    }

    setType(type = TYPE_STRING, formatCode = FORMAT_CODE_GENERAL) {
        this.type = type
        this.formatCode = formatCode
    }

    fontsAreEquals(cellStyle) {
        return fontsAreEquals(this.font, cellStyle.font)
    }

    alignmentsAreEquals(cellStyle) {
        return alignmentsAreEquals(this.alignment, cellStyle.alignment)
    }

    bordersAreEquals(cellStyle) {
        return bordersAreEquals(this.borderLeft, cellStyle.borderLeft)
            && bordersAreEquals(this.borderRight, cellStyle.borderRight)
            && bordersAreEquals(this.borderTop, cellStyle.borderTop)
            && bordersAreEquals(this.borderBottom, cellStyle.borderBottom)
            && bordersAreEquals(this.borderDiagonal, cellStyle.borderDiagonal)
    }

    fillsAreEquals(cellStyle) {
        return fillsAreEquals(this.fill, cellStyle.fill)
    }

    hasAlignments() {
        return this.alignment.horizontal !== HORIZONTAL_ALIGNMENT_DEFAULT
            || this.alignment.vertical !== VERTICAL_ALIGNMENT_DEFAULT
            || this.alignment.wrapText
    }

    hasBorders() {
        return this.borderLeft.thickness !== BORDER_THICKNESS_NONE
            || this.borderRight.thickness !== BORDER_THICKNESS_NONE
            || this.borderTop.thickness !== BORDER_THICKNESS_NONE
            || this.borderBottom.thickness !== BORDER_THICKNESS_NONE
            || this.borderDiagonal.thickness !== BORDER_THICKNESS_NONE
    }

    hasFills() {
        return this.fill.pattern !== FILL_PATTERN_NONE
            && this.fill.bgColor !== COLOR_DEFAULT
    }

    clone() {
        const result = new CellStyle()
        const font = this.font
        const alignment = this.alignment
        const borderLeft = this.borderLeft
        const borderRight = this.borderRight
        const borderTop = this.borderTop
        const borderBottom = this.borderBottom
        const borderDiagonal = this.borderDiagonal
        const fill = this.fill
        result.setFont(font.name, font.size, font.color, font.bold, font.italic, font.underline, font.strikethrough)
        result.setAlignment(alignment.horizontal, alignment.vertical, alignment.wrapText)
        result.setBorderLeft(borderLeft.thickness, borderLeft.color)
        result.setBorderRight(borderRight.thickness, borderRight.color)
        result.setBorderTop(borderTop.thickness, borderTop.color)
        result.setBorderBottom(borderBottom.thickness, borderBottom.color)
        result.setBorderDiagonal(borderDiagonal.thickness, borderDiagonal.color)
        result.setFill(fill.pattern, fill.bgColor)
        result.setType(this.type, this.formatCode)
        return result
    }
}

CellStyle.FONT_NAME_DEFAULT = FONT_NAME_DEFAULT
CellStyle.FONT_SIZE_DEFAULT = FONT_SIZE_DEFAULT

CellStyle.HORIZONTAL_ALIGNMENT_DEFAULT = HORIZONTAL_ALIGNMENT_DEFAULT
CellStyle.HORIZONTAL_ALIGNMENT_LEFT = HORIZONTAL_ALIGNMENT_LEFT
CellStyle.HORIZONTAL_ALIGNMENT_CENTER = HORIZONTAL_ALIGNMENT_CENTER
CellStyle.HORIZONTAL_ALIGNMENT_RIGHT = HORIZONTAL_ALIGNMENT_RIGHT

CellStyle.VERTICAL_ALIGNMENT_DEFAULT = VERTICAL_ALIGNMENT_DEFAULT
CellStyle.VERTICAL_ALIGNMENT_TOP = VERTICAL_ALIGNMENT_TOP
CellStyle.VERTICAL_ALIGNMENT_CENTER = VERTICAL_ALIGNMENT_CENTER
CellStyle.VERTICAL_ALIGNMENT_BOTTOM = VERTICAL_ALIGNMENT_BOTTOM

CellStyle.BORDER_THICKNESS_NONE = BORDER_THICKNESS_NONE
CellStyle.BORDER_THICKNESS_THIN = BORDER_THICKNESS_THIN

CellStyle.FILL_PATTERN_NONE = FILL_PATTERN_NONE
CellStyle.FILL_PATTERN_SOLID = FILL_PATTERN_SOLID

CellStyle.COLOR_DEFAULT = COLOR_DEFAULT

CellStyle.TYPE_STRING = TYPE_STRING
CellStyle.TYPE_NUMERIC = TYPE_NUMERIC

CellStyle.FORMAT_CODE_GENERAL = FORMAT_CODE_GENERAL

export default CellStyle
