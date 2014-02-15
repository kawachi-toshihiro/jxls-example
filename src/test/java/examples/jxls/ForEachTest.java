package examples.jxls;

import java.io.IOException;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import junit.framework.TestCase;
import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.XLSTransformer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

public class ForEachTest extends TestCase {

    public static final String rowsMergeSrcXls = "/templates/forEachRowsMerge.xls";
    public static final String rowsMergeDestXls = "target/forEachRowsMerge_output.xls";


    @Test
    public void testForEach中での行結合() throws IOException, ParsePropertyException, InvalidFormatException, URISyntaxException {
        List<HogeBean> hogeList = new ArrayList<HogeBean>();
        hogeList.add(new HogeBean("あいうえお"));
        hogeList.add(new HogeBean("abcde"));
        hogeList.add(new HogeBean("12345"));

        Map<String, Object> map = new HashMap<String, Object>();
        map.put("hogeList", hogeList);

        map.put("poiHelper", new PoiHelper());

        String path = this.getClass().getResource(rowsMergeSrcXls).getPath();

        XLSTransformer transformer = new XLSTransformer();
        transformer.transformXLS(path, map,rowsMergeDestXls);
		System.out.println("out!" + rowsMergeDestXls);

    }


}

