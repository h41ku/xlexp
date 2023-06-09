export default function downloadAs(sourceAsblob, fileName) {
    const link = document.createElement("a")
    link.href = URL.createObjectURL(sourceAsblob)
    link.download = fileName
    link.click()
    link.remove()
}
