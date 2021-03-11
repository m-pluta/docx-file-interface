public class sqlRow {
	String[] data;

	public sqlRow(String data1, String data2, String data3, String data4) {
		data[0] = data1;
		data[1] = data2;
		data[2] = data3;
		data[3] = data4;
	}

	public void setElement(int index, String input) {
		data[index] = input;
	}

	public String get(int index) {
		return data[index];

	}
}
