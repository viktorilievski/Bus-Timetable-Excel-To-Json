import service.ExcelToJson

fun main(args: Array<String>) {
    val excelToJson = ExcelToJson()
    val filePath = "./relations.xlsx"
    excelToJson.read(filePath)
}