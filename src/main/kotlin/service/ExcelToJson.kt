package service

import model.BusRelation
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.*
import kotlin.math.ceil


class ExcelToJson() {

    fun read(filePath: String) {
        val inputStream = FileInputStream(filePath)
        val workBook = WorkbookFactory.create(inputStream)

        val sheet = workBook.getSheetAt(0)
        val out: Writer = BufferedWriter(OutputStreamWriter(FileOutputStream("excelToJsonConverted.json"), "UTF-8"))
        out.write("{\"relacii\": [")
        var relationCounter = 0
        for (currentRowIterator in FIRST_ROW..LAST_ROW) {
            val row = sheet.getRow(currentRowIterator)
            if (row.getCell(FIRST_COLUMN) != null &&
                row.getCell(FIRST_COLUMN).toString().isNotEmpty() &&
                row.getCell(5).toString() != "/"
            ) {
                out.write(readDepartureRelations(relationCounter, sheet, currentRowIterator))
                relationCounter += row.getCell(CELL_OF_NUMBER_DEPARTURE_RELATIONS).numericCellValue.toInt()
                out.write(readReturningRelations(relationCounter, sheet, currentRowIterator))
                relationCounter += row.getCell(CELL_OF_NUMBER_RETURN_RELATIONS).numericCellValue.toInt()
            }
        }
        out.write("]\n}")
        out.close()
    }

    private fun readReturningRelations(currentRelationCounter: Int, sheet: Sheet, currentRowIterator: Int): String {
        var relationCounter = currentRelationCounter
        var currentRow = currentRowIterator
        var row = sheet.getRow(currentRow)
        var numberOfReturningRelations = 0
        try {
            numberOfReturningRelations = row.getCell(CELL_OF_NUMBER_RETURN_RELATIONS).numericCellValue.toInt()
        } catch (e: Exception) {
            println(row.rowNum)
            e.printStackTrace()
        }
        val companyName = row.getCell(FIRST_COLUMN).toString()
        val startPoint = row.getCell(2).toString().split(" - ")[0].trim()
        val endPoint = row.getCell(2).toString().split(" - ")[1].trim()
        var departureTime: String = ""
        var note = row.getCell(LAST_COLUMN).toString()
        var departureRelationsJson = ""
        if (row.getCell(CELL_OF_NUMBER_RETURN_RELATIONS).numericCellValue.toInt() <= 10) {
            try {
                while (row.getCell(STATION_NAME).toString().trim() != endPoint.trim()) {
                    currentRow++
                    row = sheet.getRow(currentRow)
                }
            } catch (e: Exception) {
                println(row.rowNum)
                println(row.getCell(STATION_NAME))
                println(row.getCell(17).toString())
                e.printStackTrace()
            }

            for (i in 17 until numberOfReturningRelations + 17) {
                var arrivalTime = getArrivalTime(currentRow = currentRow, sheet = sheet, cellPosition = i, isDeparture = false)
                if (row.getCell(i).toString() == "/") {
                    continue
                }
                if (row.getCell(i).toString().length > 5) { // special time and note
                    departureTime = getSpecialDepartureTime(row, i)
                    note = getSpecialNote(row, i)
                } else {
                    try {
                        val departureTimeH = row.getCell(i).toString().split(".")[0].padStart(2, '0')
                        val departureTimeM = row.getCell(i).toString().split(".")[1].padEnd(2, '0')
                        departureTime = "$departureTimeH:$departureTimeM"
                    } catch (e: Exception) {
                        print(row.rowNum)
                    }
                }
                relationCounter++
                val busRelation = BusRelation(
                    relationCounter,
                    companyName,
                    endPoint,
                    startPoint,
                    departureTime,
                    arrivalTime,
                    note
                )
                departureRelationsJson += busRelation.toJson()
            }
        } else {
            val numOfCycles = ceil(numberOfReturningRelations.toDouble() / 10).toInt()
            for (i in 0 until numOfCycles) {
                try {
                    while (sheet.getRow(currentRow + 1).getCell(16).toString() != endPoint) {
                        currentRow++
                    }
                    currentRow++
                    row = sheet.getRow(currentRow)
                } catch (e: Exception) {
                    println(row.rowNum)
                    println(endPoint)
                    e.printStackTrace()
                }

                for (j in 17 until 27) {
                    if (row.getCell(j) == null || row.getCell(j).toString() == "")
                        break
                    val arrivalTime = getArrivalTime(currentRow = currentRow, sheet = sheet, cellPosition = j, isDeparture = false)
                    if (row.getCell(j).toString() == "/") {
                        continue
                    }
                    if (row.getCell(j).toString().length > 5) { // special time and note
                        departureTime = getSpecialDepartureTime(row, j)
                        note = getSpecialNote(row, j)
                    } else {
                        val departureTimeH = row.getCell(j).toString().split(".")[0].padStart(2, '0')
                        val departureTimeM = row.getCell(j).toString().split(".")[1].padEnd(2, '0')
                        departureTime = "$departureTimeH:$departureTimeM"
                    }
                    relationCounter++
                    val busRelation = BusRelation(
                        relationCounter,
                        companyName,
                        endPoint,
                        startPoint,
                        departureTime,
                        arrivalTime,
                        note
                    )
                    departureRelationsJson += busRelation.toJson()
                }
            }
        }
        return departureRelationsJson
    }

    private fun readDepartureRelations(currentRelationCounter: Int, sheet: Sheet, currentRowIterator: Int): String {
        var relationCounter = currentRelationCounter
        var currentRow = currentRowIterator
        var row = sheet.getRow(currentRow)
        val numberOfDepartureRelations = row.getCell(CELL_OF_NUMBER_DEPARTURE_RELATIONS).numericCellValue.toInt()
        val companyName = row.getCell(FIRST_COLUMN).toString()
        val startPoint = row.getCell(2).toString().split(" - ")[0].trim()
        val endPoint = row.getCell(2).toString().split(" - ")[1].trim()
        var departureTime: String = ""
        var note = row.getCell(LAST_COLUMN).toString()
        var departureRelationsJson = ""
        if (row.getCell(CELL_OF_NUMBER_DEPARTURE_RELATIONS).numericCellValue.toInt() <= 10) {
            for (i in 5 until numberOfDepartureRelations + 5) {
                var arrivalTime = getArrivalTime(currentRow = currentRow, sheet = sheet, cellPosition = i, isDeparture = true)

                if (row.getCell(i).toString() == "/") {
                    continue
                }
                if (row.getCell(i).toString().length > 5) { // special time and note
                    departureTime = getSpecialDepartureTime(row, i)
                    note = getSpecialNote(row, i)
                } else {
                    val departureTimeH = row.getCell(i).toString().split(".")[0].padStart(2, '0')
                    val departureTimeM = row.getCell(i).toString().split(".")[1].padEnd(2, '0')
                    departureTime = "$departureTimeH:$departureTimeM"
                }
                relationCounter++
                val busRelation = BusRelation(
                    relationCounter,
                    companyName,
                    startPoint,
                    endPoint,
                    departureTime,
                    arrivalTime,
                    note
                )
                departureRelationsJson += busRelation.toJson()
            }
        } else {
            val numOfCycles = ceil(numberOfDepartureRelations.toDouble() / 10).toInt()
            for (i in 0 until numOfCycles) {
                if (i >= 1) {
                    while (row.getCell(5).toString() != "") {
                        currentRow++
                        row = sheet.getRow(currentRow)
                    }
                    currentRow++
                    row = sheet.getRow(currentRow)
                }
                for (j in 0 until 10) {

                    if (row.getCell(j + 5) == null || row.getCell(j + 5).toString() == "")
                        break
                    var arrivalTime = getArrivalTime(currentRow = currentRow, sheet = sheet, cellPosition = j + 5, isDeparture = true)
                    if (row.getCell(j + 5).toString() == "/") {
                        continue
                    }
                    if (row.getCell(j + 5).toString().length > 5) { // special time and note
                        departureTime = getSpecialDepartureTime(row, j + 5)
                        note = getSpecialNote(row, j + 5)
                    } else {
                        val departureTimeH = row.getCell(j + 5).toString().split(".")[0].padStart(2, '0')
                        val departureTimeM = row.getCell(j + 5).toString().split(".")[1].padEnd(2, '0')
                        departureTime = "$departureTimeH:$departureTimeM"
                    }
                    relationCounter++
                    val busRelation = BusRelation(
                        relationCounter,
                        companyName,
                        startPoint,
                        endPoint,
                        departureTime,
                        arrivalTime,
                        note
                    )
                    departureRelationsJson += busRelation.toJson()
                }
            }
        }

        return departureRelationsJson
    }

    private fun getArrivalTime(currentRow: Int, sheet: Sheet, cellPosition: Int, isDeparture: Boolean): String {
        var arrivalRow = currentRow
        while (sheet.getRow(arrivalRow).getCell(cellPosition) != null &&
            sheet.getRow(arrivalRow).getCell(cellPosition).toString().isNotEmpty()
        ) {
                if (isDeparture) {
                    arrivalRow++
                } else {
                    arrivalRow--
                }
        }
        println(arrivalRow)
        if (isDeparture) {
            arrivalRow--
        } else {
            arrivalRow++
        }
        val arrivalTimeH =
            sheet.getRow(arrivalRow).getCell(cellPosition).toString().split(".")[0].padStart(2, '0')
        val arrivalTimeM =
            sheet.getRow(arrivalRow).getCell(cellPosition).toString().split(".")[1].padEnd(2, '0')

        return "$arrivalTimeH:$arrivalTimeM"
    }

    private fun getSpecialNote(row: Row, position: Int): String {
        try {
            return row.getCell(position).toString().split(" - ")[1]
        } catch (e: Exception) {
            println(row.rowNum)
            e.printStackTrace()
        }
        return ""
    }

    private fun getSpecialDepartureTime(row: Row, position: Int): String {
        val time = row.getCell(position).toString().split(" - ")[0]
        try {
            return time.split(".")[0].padStart(2, '0') + ":" + time.split(".")[1].padEnd(2, '0')
        } catch (e: java.lang.Exception) {
            println(row.rowNum)
            println(time)
            e.printStackTrace()
        }
        return ""
    }

    companion object {
        private const val FIRST_ROW = 8
        private const val LAST_ROW = 3827
        private const val FIRST_COLUMN = 0
        private const val LAST_COLUMN = 27
        private const val CELL_OF_NUMBER_DEPARTURE_RELATIONS = 3
        private const val CELL_OF_NUMBER_RETURN_RELATIONS = 4
        private const val STATION_NAME = 16

    }
}