package com.kbn.excel;

import org.robotframework.javalib.library.AnnotationLibrary;
import org.robotframework.javalib.library.KeywordDocumentationRepository;
import org.robotframework.javalib.library.RobotJavaLibrary;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;

public class JyExcelLibrary implements KeywordDocumentationRepository, RobotJavaLibrary{
    public static final String ROBOT_LIBRARY_SCOPE = "GLOBAL";
    public static JyExcelLibrary instance;
    private final AnnotationLibrary annotationLibrary = new AnnotationLibrary(
            "com/kbn/excel/keyword/**.class");
    private static final String LIBRARY_DOCUMENTATION = "JyExcelLibrary is a Robot Framework test library for testing Java Swing user interfaces.\n\n"
            + "It uses a tool called [http://java.net/projects/jemmy/|Jemmy] internally to operate on Swing components.\n"
            + "= Getting Started =\n"
            + "First, the JyExcelLibrary needs to be taken into use in the settings table:\n"
            + "| *Settings * | *Value* |\n"
            + "| Library | JyExcelLibrary |\n"

            + "The tested application can be started with keyword `Start Application`, using the name "
            + "of the main application class as an argument:\n"
            + "| `Start Application` | com.acme.TheApplication |\n"
            + "| `Select Window`     | TheApplication Window   |\n"
            + "| `Push Button`       | AcmeButton              |\n"
            + "When the tests are executed, both the JyExcelLibrary and the application and all its dependencies "
            + "need to be available in the CLASSPATH. Robot Framework needs to be started with `jybot` start script "
            + "when using the JyExcelLibrary. In Windows, this can be done like:\n"
            + "| set CLASSPATH=JyExcelLibrary-<version>.jar;myApp.jar\n"
            + "| jybot my_test.txt\n"
            + "and in *nix like this:\n"
            + "| CLASSPATH=JyExcelLibrary-<version>.jar:myApp.jar jybot my_test.txt\n"

            + "= Contexts =\n"
            + "Keywords that operate on a component always search for the component in some context, "
            + "which has to explicitly set. "
            + "Allowed contexts are windows, dialogs, internal frames, and tabbed panes. "
            + "After a context has been selected, all subsequent keywords search for components in that context "
            + "until a new context is selected. Keywords that can be used to select a context are "
            + "`Select Window`, `Select Dialog` and `Select Context`. For example:\n"
            + "| `Select Window` | My App |\n"
            + "| `Select From Main Menu` | File|Exit |\n"
            + "| `Select Dialog` | Confirm |\n"
            + "| `Push Button`   | No      |\n"
            + "| `Select Window` | My App  |\n"

            + "= Locating components =\n"
            + "Most of the keywords that operate on a visible component take an argument named `identifier`, "
            + "which is used to locate the element. The first matching element is operated on, according to these rules:\n"
            + "- If the `identifier` is a number, it is used as a zero-based index for the particular component type in "
            + "the current context. Using indices is, however, fragile and is strongly discouraged.\n"
            + "- If the `identifier` matches to internal name of a component (set using `setName` method in Java code), that component is chosen.\n"
            + "- For components that have visible text (e.g. buttons), `identifier` is also matched against that.\n"
            + "- Text field keywords also support accessing awt-text fields by prefixing the identifier with awt=.\n"
            + "keyword `List Components in Context` lists all components and their names and indices in a given context.\n"

            + "= Running keywords in separate threads =\n"
            + "Some actions may cause dialogs or other components to pop up and the keyword would then not return until\n"
            + "the new dialog is closed. In these situations the keyword should be executed in a separate thread,\n"
            + "`Run keyword In Separate Thread` that test execution can continue.\n"

            + "= Logging =\n"
            + "Starting from version 1.8.0, the internal logging of Jemmy is available on Robot logs, when using DEBUG log level.";
    public JyExcelLibrary() {
        this(Collections.<String> emptyList());
    }

    protected JyExcelLibrary(final String keywordPattern) {
        this(new ArrayList<String>() {
            {
                add(keywordPattern);
            }
        });
    }

    protected JyExcelLibrary(Collection<String> keywordPatterns) {
        addKeywordPatterns(keywordPatterns);
//        disableOutput();
//        setDefaultTimeouts();
        instance = this;
    }

    private void addKeywordPatterns(Collection<String> keywordPatterns) {
        for (String pattern : keywordPatterns) {
            annotationLibrary.addKeywordPattern(pattern);
        }
    }


    public Object runKeyword(String keywordName, Object[] args) {
        return annotationLibrary.runKeyword(keywordName, toStrings(args));
    }


    public String[] getKeywordArguments(String keywordName) {
        return annotationLibrary.getKeywordArguments(keywordName);
    }


    public String getKeywordDocumentation(String keywordName) {
        if (keywordName.equals("__intro__"))
            return LIBRARY_DOCUMENTATION;
        return annotationLibrary.getKeywordDocumentation(keywordName);
    }


    public String[] getKeywordNames() {
        return annotationLibrary.getKeywordNames();
    }

//    private void setDefaultTimeouts() {
//        new TimeoutKeywords().setJemmyTimeouts("10");
//    }
//
//    private void disableOutput() {
//        TestOut out = new StandardOutOutput();
//        JemmyProperties.setCurrentOutput(out);
//    }

    private Object[] toStrings(Object[] args) {
        Object[] newArgs = new Object[args.length];
        for (int i = 0; i < newArgs.length; i++) {
            if (args[i].getClass().isArray()) {
                newArgs[i] = args[i];
            } else {
                newArgs[i] = args[i].toString();
            }
        }
        return newArgs;
    }

}
