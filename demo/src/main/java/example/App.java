package example;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;

import java.io.InputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class App {
    public static void putPictureCentered(Sheet sheet, String picturePath, int pictureType, int colIdx, int rowIdx) throws Exception {
        Workbook wb = sheet.getWorkbook();

        // load the picture
        InputStream inputStream = new FileInputStream(picturePath);
        byte[] bytes = IOUtils.toByteArray(inputStream);
        int pictureIdx = wb.addPicture(bytes, pictureType);
        inputStream.close();

        // create an anchor with upper left cell colIdx/rowIdx, only one cell anchor since bottom right depends on resizing
        CreationHelper helper = wb.getCreationHelper();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(colIdx);
        anchor.setRow1(rowIdx);

        // create a picture anchored to colIdx and rowIdx
        Drawing<?> drawing = (Drawing<?>) sheet.createDrawingPatriarch();
        Picture pict = drawing.createPicture(anchor, pictureIdx);

        // get the picture width in px
        int pictWidthPx = pict.getImageDimension().width;
        // get the picture height in px
        int pictHeightPx = pict.getImageDimension().height;

        // get column width of column in px
        float columnWidthPx = sheet.getColumnWidthInPixels(colIdx);

        // get the height of row in px
        Row row = sheet.getRow(rowIdx);
        float rowHeightPt = row.getHeightInPoints();
        float rowHeightPx = rowHeightPt * Units.PIXEL_DPI / Units.POINT_DPI;

        // is horizontal centering possible?
        if (pictWidthPx <= columnWidthPx) {

            // calculate the horizontal center position
            int horCenterPosPx = Math.round(columnWidthPx / 2f - pictWidthPx / 2f);
            // set the horizontal center position as Dx1 of anchor
            if (wb instanceof XSSFWorkbook) {
                anchor.setDx1(horCenterPosPx * Units.EMU_PER_PIXEL); //in unit EMU for XSSF
            } else if (wb instanceof HSSFWorkbook) {
                // see https://stackoverflow.com/questions/48567203/apache-poi-xssfclientanchor-not-positioning-picture-with-respect-to-dx1-dy1-dx/48607117#48607117 for HSSF
                int DEFAULT_COL_WIDTH = 10 * 256;
                anchor.setDx1(Math.round(horCenterPosPx * Units.DEFAULT_CHARACTER_WIDTH / 256f * 14.75f * DEFAULT_COL_WIDTH / columnWidthPx));
            }

        } else {
            System.out.println("Picture is too width. Horizontal centering is not possible.");
        }

        // is vertical centering possible?
        if (pictHeightPx <= rowHeightPx) {

            // calculate the vertical center position
            int vertCenterPosPx = Math.round(rowHeightPx / 2f - pictHeightPx / 2f);
            // set the vertical center position as Dy1 of anchor
            if (wb instanceof XSSFWorkbook) {
                anchor.setDy1(Math.round(vertCenterPosPx * Units.EMU_PER_PIXEL)); //in unit EMU for XSSF
            } else if (wb instanceof HSSFWorkbook) {
                // see https://stackoverflow.com/questions/48567203/apache-poi-xssfclientanchor-not-positioning-picture-with-respect-to-dx1-dy1-dx/48607117#48607117 for HSSF
                float DEFAULT_ROW_HEIGHT = 12.75f;
                anchor.setDy1(Math.round(vertCenterPosPx * Units.PIXEL_DPI / Units.POINT_DPI * 14.75f * DEFAULT_ROW_HEIGHT / rowHeightPx));
            }

        } else {
            System.out.println("Picture is too height. Vertical centering is not possible.");
        }

        // resize the picture to it's native size
        pict.resize();
    }

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();
        String resultName = "CenterImageTest.xlsx";
        Sheet sheet = wb.createSheet("Sheet1");

        int colIdx = 1; // cell B
        int colWidth = 20; // in default character widths
        int rowIdx = 1; // row 2
        float rowHeight = 200; // in points

        //========================prepare sheet
        // create cell
        Row row = sheet.createRow(rowIdx);
        Cell cell = row.createCell(colIdx);
        // set column width of colIdx in default character widths
        sheet.setColumnWidth(colIdx, colWidth * 256);
        // set row height of rowIdx in points
        row.setHeightInPoints(rowHeight);
        //========================end prepare sheet

        // put image centered
        String picturePath = "./pict100x100.png"; // small image
        // String picturePath = "./pict100x200.png"; // image too height
        // String picturePath = "./pict200x100.png"; // image too width
        // String picturePath = "./pict200x200.png"; // image too bir

        try {
            putPictureCentered(sheet, picturePath, Workbook.PICTURE_TYPE_PNG, colIdx, rowIdx);
            FileOutputStream fileOut = new FileOutputStream("./" + resultName);
            wb.write(fileOut);
            fileOut.close();
            wb.close();
        } catch (Exception e) {
            System.out.println(e);
        }
    }
}
