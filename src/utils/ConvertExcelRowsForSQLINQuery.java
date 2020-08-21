package utils;

public class AddCommaInExcelCells {
    private static final String EXCEL_ISO_CODES = "POCOND1\n" +
            "POCOND2\n" +
            "POCOND3\n" +
            "ENTCVMI4\n" +
            "ENTCVMI3\n" +
            "ENTCVMI2\n" +
            "ENTCVMI1\n" +
            "POCNMD1\n" +
            "POCNMD2\n" +
            "POCNMD3\n" +
            "IT050\n" +
            "IT100\n" +
            "IT200\n" +
            "IT300\n";
    public static void main(String[] args) {
        String isoCodesWithComma = EXCEL_ISO_CODES.replace("\n",",");
        String[] split = isoCodesWithComma.split(",");
        StringBuilder sb = new StringBuilder();
        for(String each : split) {
            if(null != each) {
                String newEach = "'"+each+"',";
                sb.append(newEach);
            }

        }
        System.out.println(sb.toString());
    }
}
