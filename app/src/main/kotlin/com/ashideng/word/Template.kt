package com.ashideng.word

import com.spire.doc.*
import com.spire.doc.documents.HorizontalAlignment
import com.spire.doc.documents.Paragraph
import com.spire.doc.documents.TableRowHeightType
import com.spire.doc.documents.VerticalAlignment
import com.spire.doc.fields.TextRange
import java.awt.Color


class Template {

    fun generteTemplate() {
        val doc = Document()
        val section = doc.addSection()

        val header = listOf("Name", "Capital", "Continent", "Area", "Population")
        val data = listOf(
            listOf("Argentina", "Buenos Aires", "South America", "2777815", "32300003"),
            listOf("Bolivia", "La Paz", "South America", "1098575", "7300000"),
            listOf("Brazil", "Brasilia", "South America", "8511196", "150400000"),
            listOf("Canada", "Ottawa", "North America", "9976147", "26500000"),
            listOf("Chile", "Santiago", "South America", "756943", "13200000"),
            listOf("Colombia", "Bogota", "South America", "1138907", "33000000"),
            listOf("Cuba", "Havana", "North America", "114524", "10600000"),
            listOf("Ecuador", "Quito", "South America", "455502", "10600000"),
            listOf("El Salvador", "San Salvador", "North America", "20865", "5300000"),
            listOf("Guyana", "Georgetown", "South America", "214969", "800000")
        )

        val table = section.addTable(true)
        table.resetCells(data.size+1, header.size)

        val firstRow = table.rows.first() as TableRow
        with(firstRow) {
            isHeader(true)
            height=20F
            heightType= TableRowHeightType.Exactly
            rowFormat.backColor= Color.gray
        }

        for (i in 0 until header.size) {
            firstRow.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle)
            val p: Paragraph = firstRow.getCells().get(i).addParagraph()
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center)
            val txtRange: TextRange = p.appendText(header[i])
            txtRange.getCharacterFormat().setBold(true)
        }

        //Add data to the rest of rows

        //Add data to the rest of rows
        for (r in 0 until data.size) {
            val dataRow = table.rows[r + 1]
            dataRow.setHeight(25F)
            dataRow.heightType = TableRowHeightType.Exactly
            dataRow.rowFormat.backColor = Color.white
            for (c in 0 until data[r].size) {
                dataRow.cells[c].cellFormat.verticalAlignment = VerticalAlignment.Middle
                dataRow.cells[c].addParagraph().appendText(data[r][c])
            }
        }

        //Set background color for cells

        //Set background color for cells
        for (j in 1 until table.rows.count) {
            if (j % 2 == 0) {
                val row2 = table.rows[j]
                for (f in 0 until row2.cells.count) {
                    row2.cells[f].cellFormat.backColor = Color(173, 216, 230)
                }
            }
        }

        doc.saveToFile("output/CreateTable.docx", FileFormat.Docx_2013);

    }
}