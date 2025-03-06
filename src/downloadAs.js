export default function downloadAs(sourceAsBlob, fileName) {
    const link = document.createElement("a")
    link.href = URL.createObjectURL(sourceAsBlob)
    link.download = fileName
    link.click()
    link.remove()
}
