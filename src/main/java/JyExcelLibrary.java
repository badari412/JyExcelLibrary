import java.util.Collection;
import java.util.Collections;

public class JyExcelLibrary extends com.kbn.excel.JyExcelLibrary {

    public JyExcelLibrary() {
        this(Collections.<String>emptyList());
    }

    protected JyExcelLibrary(Collection<String> keywordPatterns) {
        super(keywordPatterns);
    }
}
