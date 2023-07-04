package com.newcubator.jdbcexcel;

import static java.util.Collections.emptyMap;

import com.newcubator.jdbcexcel.cellwriters.BigDecimalCellWriter;
import com.newcubator.jdbcexcel.cellwriters.BooleanCellWriter;
import com.newcubator.jdbcexcel.cellwriters.DateCellWriter;
import com.newcubator.jdbcexcel.cellwriters.HintTextCellWriter;
import com.newcubator.jdbcexcel.cellwriters.NumberCellWriter;
import com.newcubator.jdbcexcel.cellwriters.ObjectCellWriter;
import com.newcubator.jdbcexcel.cellwriters.ReplaceableStringCellWriter;
import com.newcubator.jdbcexcel.cellwriters.StringCellWriter;
import com.newcubator.jdbcexcel.configuration.ExportConfiguration;
import com.newcubator.jdbcexcel.tabs.ExcelTab;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Date;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.jdbc.core.ArgumentPreparedStatementSetter;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowCallbackHandler;
import org.springframework.lang.NonNull;
import org.springframework.util.StringUtils;

@Slf4j
public class ExcelWriter {

    public static final int ROWS_IN_MEMORY = 500;

    // More then 100 characters are probably not necessary
    // Results in a column width of 114
    public static final int MAX_COLUMN_WIDTH_CHARS = 100;

    protected final JdbcTemplate template;
    protected final Map<String, String> stringReplacements;

    public ExcelWriter(JdbcTemplate template) {
        this(template, emptyMap());
    }

    public ExcelWriter(JdbcTemplate template, Map<String, String> replacements) {
        this.template = template;
        this.stringReplacements = replacements;
    }

    public byte[] createExcel(ExcelTab exportTab) throws IOException {
        return createExcel(List.of(exportTab), new ExportConfiguration());
    }

    public byte[] createExcel(ExcelTab exportTab, ExportConfiguration exportConfiguration) throws IOException {
        return createExcel(List.of(exportTab), exportConfiguration);
    }

    public byte[] createExcel(List<ExcelTab> exportTabs) throws IOException {
        return this.createExcel(exportTabs, new ExportConfiguration());
    }

    public byte[] createExcel(List<ExcelTab> exportTabs, ExportConfiguration exportConfiguration) throws IOException {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(ROWS_IN_MEMORY)) {
            exportTabs.forEach((tab) -> {
                log.debug("Adding sheet {}", tab.getName());

                SXSSFSheet fieldSheet = workbook.createSheet(tab.getName());

                List<Object> arguments = tab.getParameter();
                String sql = tab.getSql();

                if (checkParameterCount(sql, arguments.size())) {
                    log.debug("Adding {} as parameters to query", arguments);
                    addTab(workbook, fieldSheet, sql, arguments.toArray(), tab.getHintText(), exportConfiguration);
                } else {
                    log.warn("Unable to add sheet {} cause {} are required but {} given", tab, arguments, arguments.size());
                }
            });

            return createByteArray(workbook);
        }
    }

    private boolean checkParameterCount(String sqlStatement, int actualParameterCount) {
        int paramCount = StringUtils.countOccurrencesOf(sqlStatement, "?");
        return paramCount == actualParameterCount;
    }

    private byte[] createByteArray(SXSSFWorkbook workbook) throws IOException {
        try (ByteArrayOutputStream stream = new ByteArrayOutputStream(1_000_000)) {
            workbook.write(stream);

            workbook.dispose();

            return stream.toByteArray();
        }
    }

    private void addTab(SXSSFWorkbook workbook,
                        SXSSFSheet exportSheet,
                        String sql,
                        Object[] arguments,
                        String hintText,
                        ExportConfiguration exportConfiguration) {
        log.info("Adding tab for query '{}' to export", sql);
        template.query(
            sql,
            new ArgumentPreparedStatementSetter(arguments),
            new ExportCallbackHandler(workbook, exportSheet, stringReplacements, hintText, exportConfiguration)
        );
    }

    static class ExportCallbackHandler implements RowCallbackHandler {

        private final SXSSFWorkbook workbook;
        private final SXSSFSheet exportSheet;

        private final DateCellWriter dateCellWriter;
        private final StringCellWriter stringCellWriter;
        private final ReplaceableStringCellWriter replaceableStringCellWriter;
        private final NumberCellWriter numberCellWriter;
        private final BigDecimalCellWriter bigDecimalCellWriter;
        private final ObjectCellWriter objectCellWriter;
        private final BooleanCellWriter booleanCellWriter;

        private int[] maxColumnWidths;

        // Indicates at which row index the data table with database data starts
        private int dataRowsStartIndex = 0;

        ExportCallbackHandler(SXSSFWorkbook workbook,
                              SXSSFSheet exportSheet,
                              Map<String, String> replacements,
                              String hintText,
                              ExportConfiguration exportConfiguration) {
            this.workbook = workbook;
            this.exportSheet = exportSheet;

            this.dateCellWriter = new DateCellWriter();
            this.stringCellWriter =
                new StringCellWriter(exportConfiguration.isAutogenerateHyperlinks());
            this.replaceableStringCellWriter =
                new ReplaceableStringCellWriter(replacements, exportConfiguration.isAutogenerateHyperlinks());
            this.numberCellWriter = new NumberCellWriter();
            this.bigDecimalCellWriter = new BigDecimalCellWriter();
            this.objectCellWriter = new ObjectCellWriter();
            this.booleanCellWriter = new BooleanCellWriter();

            if (hintText != null) {
                writeHintTextRow(hintText);
            }
        }

        private void writeHintTextRow(String hintText) {
            SXSSFRow excelRow = exportSheet.createRow(dataRowsStartIndex);

            new HintTextCellWriter().writeCellValue(workbook, excelRow, 0, hintText);

            // Columns without data in the next column simply overflow in excel
            // We only write into the first column, so we can ignore column widths completely for this row
            setWrittenColumnWidth(0, 0);
            advanceDataRowsStartBy(2);
        }

        private void writeHeaderRow(ResultSetMetaData metaData) throws
            SQLException {

            SXSSFRow headerRow = exportSheet.createRow(dataRowsStartIndex);

            for (int i = 1; i <= metaData.getColumnCount(); i++) {
                String columnHeader = metaData.getColumnLabel(i);

                int writtenHeaderColumnWidth = stringCellWriter
                    .writeCellValue(workbook, headerRow, i - 1, columnHeader);

                setWrittenColumnWidth(i - 1, writtenHeaderColumnWidth);
            }

            advanceDataRowsStartBy(1);
        }

        @Override
        public void processRow(@NonNull ResultSet rs) throws SQLException {
            ResultSetMetaData metaData = rs.getMetaData();

            int databaseRowIndex = rs.getRow();
            if (databaseRowIndex == 1) {
                writeHeaderRow(metaData);
            }

            writeRow(rs, metaData);

            if (rs.isLast()) {
                resizeRows(exportSheet, maxColumnWidths);
            }
        }

        private void writeRow(ResultSet rs, ResultSetMetaData metaData) throws SQLException {
            // Jdbc ResultSet row indexes are 1-based, while excel rows are 0-based
            int excelRowIndex = dataRowsStartIndex + (rs.getRow() - 1);
            SXSSFRow excelRow = exportSheet.createRow(excelRowIndex);

            log.debug("Writing database tuple index {} into excel row {}", rs.getRow(), excelRowIndex);

            for (int i = 1; i <= metaData.getColumnCount(); i++) {
                int writtenCellValueLength = writeCellValueWithWriter(rs, metaData, excelRow, i);
                setWrittenColumnWidth(i - 1, writtenCellValueLength);
            }
        }

        /**
         * Writes a cell value identified by the current row and columnIndex to the excel sheet
         * and returns the length of the written value in characters
         */
        private int writeCellValueWithWriter(ResultSet rs, ResultSetMetaData metaData, SXSSFRow excelRow, int columnIndex) throws
            SQLException {

            switch (metaData.getColumnType(columnIndex)) {
                case Types.NULL:
                    return 0;

                case Types.VARCHAR:
                case Types.CHAR:
                case Types.LONGVARCHAR:
                case Types.LONGNVARCHAR:
                    String stringCellValue = rs.getString(columnIndex);

                    if (!rs.wasNull()) {
                        return replaceableStringCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            stringCellValue
                        );
                    }
                    return 0;

                case Types.DATE:
                case Types.TIMESTAMP:
                case Types.TIMESTAMP_WITH_TIMEZONE:
                case Types.TIME:
                case Types.TIME_WITH_TIMEZONE:
                    Date dateCellValue = rs.getDate(columnIndex);

                    if (!rs.wasNull()) {
                        return dateCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            dateCellValue
                        );
                    }
                    return 0;

                case Types.DOUBLE:
                case Types.INTEGER:
                case Types.SMALLINT:
                case Types.DECIMAL:
                case Types.FLOAT:
                case Types.TINYINT:
                    double doubleCellValue = rs.getDouble(columnIndex);

                    if (!rs.wasNull()) {
                        return numberCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            doubleCellValue
                        );
                    }
                    return 0;

                case Types.BIGINT:
                case Types.NUMERIC:
                    BigDecimal bigDecimalCellValue = rs.getBigDecimal(columnIndex);

                    if (!rs.wasNull()) {
                        return bigDecimalCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            bigDecimalCellValue
                        );
                    }
                    return 0;

                case Types.BOOLEAN:
                case Types.BIT:
                    boolean booleanCellValue = rs.getBoolean(columnIndex);

                    if (!rs.wasNull()) {
                        return booleanCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            booleanCellValue
                        );
                    }
                    return 0;

                case Types.OTHER:
                default:
                    Object objectCellValue = rs.getObject(columnIndex);

                    if (!rs.wasNull()) {
                        return objectCellWriter.writeCellValue(
                            workbook,
                            excelRow,
                            columnIndex - 1,
                            objectCellValue
                        );
                    }
                    return 0;
            }
        }

        private void setWrittenColumnWidth(int columnIndex, int columnWidth) {
            if (maxColumnWidths == null) {
                maxColumnWidths = new int[columnIndex + 1];
            }

            if (columnIndex >= maxColumnWidths.length) {
                maxColumnWidths = Arrays.copyOf(maxColumnWidths, columnIndex + 1);
            }

            if (maxColumnWidths[columnIndex] < columnWidth) {
                maxColumnWidths[columnIndex] = columnWidth;
            }
        }

        private void advanceDataRowsStartBy(int rows) {
            dataRowsStartIndex += rows;
        }

        private void resizeRows(SXSSFSheet exportSheet, int[] maxColumnWidths) {
            int maxColumnWidth = calculateColumnWidth(MAX_COLUMN_WIDTH_CHARS);

            for (int i = 0; i < maxColumnWidths.length; i++) {
                int columnWidth = Math.min(
                    maxColumnWidth,
                    calculateColumnWidth(maxColumnWidths[i])
                );

                exportSheet.setColumnWidth(i, columnWidth);
            }
        }

        /**
         * Resize columns based on an estimated font width.
         * Since Apache poi uses font measuring (ew), autotracking colun widths gets increasingly
         * expensive if the count of rows grows that the measuring is based on. Additionally, we now have an
         * implicit dependency on fonts being available to base our measurement on.
         * This little formula just estimates the column width, which is good enough for our usecases.
         * <p>
         * https://stackoverflow.com/questions/18983203/how-to-speed-up-autosizing-columns-in-apache-poi#answer-19007294
         */
        private int calculateColumnWidth(int charCount) {
            return ((int) (charCount * 1.14388)) * 256;
        }
    }
}
