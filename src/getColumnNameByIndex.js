export default function getColumnNameByIndex(n) {

    const ordA = 'A'.charCodeAt(0)
    const ordZ = 'Z'.charCodeAt(0)
    const len = ordZ - ordA + 1

    let s = ''
    while (n >= 0) {
        s = String.fromCharCode(n % len + ordA) + s
        n = Math.floor(n / len) - 1
    }
    return s
}
