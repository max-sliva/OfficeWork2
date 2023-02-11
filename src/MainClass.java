
public class MainClass {
	public static void main(String[] args) {
		String curPath = System.getProperty("user.dir");
		System.out.println("Working Directory = " + curPath);
		DBClass myDB = new DBClass(curPath);
//		myDB.showTable("student");
//		myDB.showTable("lecturer");
//		myDB.showTable("subject");
//		myDB.showTable("university");
//		myDB.showTable("exam_marks");
//		myDB.showTable("subj_lect");
//		myDB.executeQuery("select student.SURNAME, student.NAME, student.STIPEND, student.KURS, "
//				+ "student.CITY, student.BIRTHDAY, university.univ_name from student, university where student.univ_id=university.univ_id");
		myDB.executeQuery("select lecturer.SURNAME, lecturer.NAME, lecturer.CITY, "
				+ "university.univ_name from lecturer, university where lecturer.univ_id=university.univ_id");
	}
}
