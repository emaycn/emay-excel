package cn.emay.excel.common.schema.base;

/**
 * 表定义的读参数集
 *
 * @author Frank
 */
public class SheetReadSchemaParams {

    /**
     * 标题行号<br/>
     * <p>
     * - reader:从哪一行读取标题，如果小于0则不读取标题，默认为0。【如果readColumnBy=Title，则不可以小于0】<br/>
     */
    private int readTitleRowIndex = 0;

    /**
     * 开始读取数据的行号(从0开始)<br/>
     * 当readColumnBy=Title时，数据行号必须比title号行要大<br/>
     * <p>
     * - reader:从哪一行开始读取，默认index=1<br/>
     */
    private int readDataStartRowIndex = 1;

    /**
     * 结束读取数据的行号(从0开始)<br/>
     * 当readColumnBy=Title时，数据行号必须比title号行要大<br/>
     * <p>
     * - reader:读取到哪一行结束，如果小于0则全部读取，默认-1<br/>
     */
    private int readDataEndRowIndex = -1;

    /**
     * 读取列规则[Index,Title];<br/>
     * <p>
     * - reader:匹配数据根据，默认根据列Index进行数据匹配。如果=Ttile，【则readtitleRowIndex不可以小于0】<br/>
     */
    private String readColumnBy = "Index";

    /**
     *
     */
    public SheetReadSchemaParams() {

    }

    /**
     * @param readTitleRowIndex     从哪一行读取标题，如果小于0则不读取标题，默认为0。【如果readColumnBy=Title，则不可以小于0】
     * @param readDataStartRowIndex 从哪一行开始读取，默认index=1
     * @param readDataEndRowIndex   读取到哪一行结束，如果小于0则全部读取，默认-1
     * @param readColumnBy          匹配数据根据，默认根据列Index进行数据匹配。如果=Ttile，【则readtitleRowIndex不可以小于0】
     */
    public SheetReadSchemaParams(int readTitleRowIndex, int readDataStartRowIndex, int readDataEndRowIndex, String readColumnBy) {
        this.readTitleRowIndex = readTitleRowIndex;
        this.readDataStartRowIndex = readDataStartRowIndex;
        this.readDataEndRowIndex = readDataEndRowIndex;
        this.readColumnBy = readColumnBy;
    }

    /**
     * 是否按照列序号读取数据
     *
     * @return 是否按照列序号读取数据
     */
    public boolean readByIndex() {
        return "Index".equalsIgnoreCase(readColumnBy);
    }

    public int getReadTitleRowIndex() {
        return readTitleRowIndex;
    }

    /**
     * @param readTitleRowIndex 从哪一行读取标题，如果小于0则不读取标题，默认为0。【如果readColumnBy=Title，则不可以小于0】
     */
    public void setReadTitleRowIndex(int readTitleRowIndex) {
        this.readTitleRowIndex = readTitleRowIndex;
    }

    public int getReadDataStartRowIndex() {
        return readDataStartRowIndex;
    }

    /**
     * @param readDataStartRowIndex 从哪一行开始读取，默认index=1
     */
    public void setReadDataStartRowIndex(int readDataStartRowIndex) {
        this.readDataStartRowIndex = readDataStartRowIndex;
    }

    public int getReadDataEndRowIndex() {
        return readDataEndRowIndex;
    }

    /**
     * @param readDataEndRowIndex 读取到哪一行结束，如果小于0则全部读取，默认-1
     */
    public void setReadDataEndRowIndex(int readDataEndRowIndex) {
        this.readDataEndRowIndex = readDataEndRowIndex;
    }

    public String getReadColumnBy() {
        return readColumnBy;
    }

    /**
     * @param readColumnBy 匹配数据根据，默认根据列Index进行数据匹配。如果=Ttile，【则readtitleRowIndex不可以小于0】
     */
    public void setReadColumnBy(String readColumnBy) {
        this.readColumnBy = readColumnBy;
    }
}