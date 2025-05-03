package lu.nyo.excel.renderer;

import org.apache.commons.lang3.time.StopWatch;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.LinkedList;

import static java.util.concurrent.TimeUnit.MILLISECONDS;

public class ExcelFileRendererTest {

    @Test
    public void test01() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionSimpleTableFormat1());

        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat1"));
        stopWatch.stop();
        System.out.printf("Test case 01: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test02() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionSimpleTableFormat2());

        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat2"));
        stopWatch.stop();
        System.out.printf("Test case 02: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test03() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionSimpleTableFormat3());

        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat3"));
        stopWatch.stop();
        System.out.printf("Test case 03: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test04() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionSimpleTableFormat4());

        final long start = System.nanoTime();
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat4"));
        stopWatch.stop();
        System.out.printf("Test case 04: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test05() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionSimpleTableFormat5());

        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat5"));
        stopWatch.stop();
        System.out.printf("Test case 05: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test06() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionSimpleTableFormat6());

        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat6"));
        stopWatch.stop();
        System.out.printf("Test case 06: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test07() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionSimpleTableFormat7());
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat7"));
        stopWatch.stop();
        System.out.printf("Test case 07: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test08() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionSimpleTableFormat8());
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat8"));
        stopWatch.stop();
        System.out.printf("Test case 08: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test09() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionSimpleTableFormat9());
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat9"));
        stopWatch.stop();
        System.out.printf("Test case 09: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test10() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionSimpleTableFormat10());
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat10"));
        stopWatch.stop();
        System.out.printf("Test case 10: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test11() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionWithColSpanResolutionSimpleTableFormat());
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat11"));
        stopWatch.stop();
        System.out.printf("Test case 11: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test12() throws IOException {
        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();
        data.put("Sheet 1", DataGenerator.rowSpanResolutionSimpleTableFormat());

        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat12"));
        stopWatch.stop();
        System.out.printf("Test case 12: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

    @Test
    public void test13() throws IOException {

        LinkedHashMap<String, LinkedList<Renderable>> data = new LinkedHashMap<>();

        data.put("Sheet 1 - Test", DataGenerator.simpleTableFormatData());

        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        ExcelFileRenderer.render(data, Utils.getTestCss(), Utils.getOutputStream("testRowSpanResolutionSimpleTableFormat13"));
        stopWatch.stop();
        System.out.printf("Test case 13: %s milliseconds%n", stopWatch.getTime(MILLISECONDS));
    }

}
