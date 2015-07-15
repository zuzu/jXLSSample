import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Object collection output demo
 * 
 * Excelファイル出力用のデザインテンプレートエンジン「jXLS2.0」のデモプロジェクトです。
 * Listデータの入れ込みの方法、さらに可変領域に対するsum関数の範囲追従、個別の出力の３種類の技法のデモとなります。
 */
public class ObjectCollectionDemo {
	static Logger logger = LoggerFactory.getLogger(ObjectCollectionDemo.class);

	public static void main(String[] args) throws ParseException, IOException {
		logger.info("Running Object Collection demo");
		List<Employee> employees = generateSampleEmployeeData();
		long t1 = System.nanoTime();

		// テンプレートを開き、Area（テンプレートの変更範囲）の準備まで行う。」
		// ※Areaを使うことで全体を精査する必要がなく処理速度が上がる。
		InputStream is = ObjectCollectionDemo.class.getResourceAsStream("object_collection_template.xls");
		OutputStream os = new FileOutputStream("object_collection_output.xls");
		Transformer transformer = TransformerFactory.createTransformer(is, os);
		AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
		List<Area> xlsAreaList = areaBuilder.build();
		Area xlsArea = xlsAreaList.get(0);

		// Contextへと値をセット
		Context context = new Context();
		context.putVar("employees", employees);
		context.putVar("jpstr", "文字列挿入テスト");

		// ViewへとContextオブジェクトをセット。
		// テンプレートの変更範囲を指定。以後xlsAreaを利用する。
		// Template!A1　＝　テンプレートシートのA1を参照。
		// A1のコメントとして「jx:area(lastCell="D10")」を指定してあり、A1:D10の範囲を適用範囲としている。
		xlsArea.applyAt(new CellRef("Template!A1"), context);

		// エリア内の全ての式を再計算（セルのずれなども反映して調整）
		xlsArea.processFormulas();

		// ファイルへと書き出し。
		transformer.write();

		is.close();
		os.close();
		long t2 = System.nanoTime();
		System.out.println(t2 - t1);
	}

	private static List<Employee> generateSampleEmployeeData() throws ParseException {
		List<Employee> employees = new ArrayList<Employee>();
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
		employees.add(new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15));
		employees.add(new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25));
		employees.add(new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00));
		employees.add(new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700, 0.15));
		employees.add(new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20));
		return employees;
	}
}
