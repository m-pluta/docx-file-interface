public class sqlRow {
	String[] data = new String[4];

	public sqlRow(String data1, String data2, String data3, String data4) {
		this.data[0] = data1;
		this.data[1] = data2;
		this.data[2] = data3;
		this.data[3] = data4;
	}

	public void setElement(int index, String input) {
		this.data[index] = input;
	}

	public String get(int index) {
		return this.data[index];

	}
}
