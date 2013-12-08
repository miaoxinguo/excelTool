package foo.bar.excel.core;

public enum ColumnIndex {
    last_month_gy,
    last_month_sb,
    last_month_6s,
    last_month_ldjl,
    gy_base,
    zl,
    gy,
    aq_base,
    sb_base,
    sb,
    ss_base,
    ss,
    ldjl_base,
    ldjl,
    hlhjy;
    
    private int index;

    public int getIndex(){
        return index;
    }
    
    public void setIndex(int index){
        this.index = index;
    }
    
}
