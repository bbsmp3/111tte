// The Decompiled time:2017-03-15 12:10:13
// Decompiled home page: http://www.ludaima.cn
// Decompiler options: packimports(3) deadcode fieldsfirst ansi space 
// Source File Name:   ToExcel.java

package com.cib;

import java.io.*;
import java.util.*;

import jxl.*;
import jxl.format.*;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.*;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException;

public class ToExcel
{

    private static String pathName = ".";

    public ToExcel()
    {
    }

    public static void main(String args[])
    {
        batchToExcel(pathName);
    }

    public static void batchToExcel(String pathName)
    {
        File path = new File(pathName);
        String files[] = path.list();
        String as[];
        int j = (as = files).length;
        for (int i = 0; i < j; i++)
        {
            String fileName = as[i];
            String wholeName = (new StringBuilder(String.valueOf(pathName))).append("\\").append(fileName).toString();
            File file = new File(wholeName);
            if (file.isDirectory())
                batchToExcel(wholeName);
            else
            if (wholeName.endsWith(".txt"))
            {
                List content = readFile(wholeName);
                exportExcel(content, wholeName);
            }
        }

    }

    public static List readFile(String wholeName)
    {
        List content = new ArrayList();
        try
        {
            FileInputStream fis = new FileInputStream(wholeName);
            Scanner scan;
            String row[];
            for (scan = new Scanner(fis); scan.hasNextLine(); content.add(row))
            {
                String line = scan.nextLine();
                row = line.split("\\|");
            }

            scan.close();
        }
        catch (FileNotFoundException e)
        {
            e.printStackTrace();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        return content;
    }

    public static void exportExcel(List content, String wholeName)
    {
        try
        {
            String oFileName = (new StringBuilder(String.valueOf(wholeName.substring(0, wholeName.length() - 4)))).append(".xls").toString();
            File oFile = new File(oFileName);
            if (oFile.exists())
                oFile.delete();
            OutputStream os = new FileOutputStream(oFile);
            WritableFont wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.BLUE);
            WritableCellFormat wcf = new WritableCellFormat(wf);
            wcf.setVerticalAlignment(VerticalAlignment.CENTRE);
            wcf.setAlignment(Alignment.CENTRE);
            NumberFormat DoubleFormat = new NumberFormat("0.00");
            WritableCellFormat wcf1 = new WritableCellFormat(wf);
            wcf = new WritableCellFormat(NumberFormats.DEFAULT);
            wcf1 = new WritableCellFormat(DoubleFormat);
            WritableWorkbook wwb = Workbook.createWorkbook(os);
            int sheetCount = 0;
            int rowCount = 0;
            for (; sheetCount < content.size() / 0x10000 + 1; sheetCount++)
            {
                WritableSheet sheet = wwb.createSheet((new StringBuilder("data_")).append(sheetCount).toString(), sheetCount);
                sheet.setColumnView(0, 11);
                sheet.setColumnView(1, 11);
                sheet.setColumnView(2, 11);
                sheet.setColumnView(3, 11);
                sheet.setColumnView(4, 11);
                for (int j = 0; rowCount < content.size() && j < 0x10000; rowCount++)
                {
                    String row[] = (String[])content.get(rowCount);
                    for (int i = 0; i < row.length; i++)
                        if (row[i].matches("^-?\\d+\\.\\d+$") || row[i].matches("-\\d+$"))
                            sheet.addCell(new Number(i, j, Double.valueOf(row[i]).doubleValue(), wcf1));
                        else
                            sheet.addCell(new Label(i, j, row[i], wcf));

                    j++;
                }

            }

            wwb.write();
            wwb.close();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
        catch (RowsExceededException e)
        {
            e.printStackTrace();
        }
        catch (WriteException e)
        {
            e.printStackTrace();
        }
    }

    public static void test(String pathName)
    {
        String fileName = (new StringBuilder(String.valueOf(pathName))).append("\\fj1.xls").toString();
        File template = new File(fileName);
        try
        {
            Workbook iwb = Workbook.getWorkbook(template);
            File oFile = new File((new StringBuilder(String.valueOf(pathName))).append("\\fj1_ret.xls").toString());
            if (oFile.exists())
                oFile.delete();
            OutputStream os = new FileOutputStream(oFile);
            WritableWorkbook owb = Workbook.createWorkbook(os, iwb);
            WritableSheet sheet = owb.getSheet(0);
            Cell cell = sheet.getCell(7, 7);
            sheet.addCell(new Label(7, 7, "21sdf", cell.getCellFormat()));
            Cell cell1 = sheet.getCell(5, 43);
            WritableCellFormat wcf = new WritableCellFormat(cell1.getCellFormat());
            wcf.setAlignment(Alignment.LEFT);
            sheet.addCell(new Label(5, 43, "21sdf", wcf));
            sheet.getSettings().setPassword("111");
            sheet.getSettings().setProtected(true);
            owb.write();
            owb.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

}
