package tool;

public class FunctionPoint {
	private String element;
	private Integer l;
	private Integer a;
	private Integer h;
	private Integer s;
	Integer f1 = 0, f2 = 0, f3 = 0;
	private int total;

	public FunctionPoint() {
		super();
	}

	public FunctionPoint(String element, int l, int a, int h) {
		super();
		this.element = element;
		this.l = l;
		this.a = a;
		this.h = h;
	}

	public String getEle() {
		return element;
	}

	public void setEle(String ele) {
		this.element = ele;
	}

	public Integer getA() {
		if (a == null)
			return 0;
		return a;
	}

	public void setA(int a) {
		this.a = a;
	}

	public Integer getL() {
		if (l == null)
			return 0;
		return l;
	}

	public void setL(int l) {
		this.l = l;
	}

	public Integer getH() {
		if (h == null)
			return 0;
		return h;
	}

	public void setH(int h) {
		this.h = h;
	}

	public void setS(int s) {
		this.s = s;
	}

	public Integer getF1() {
		return f1;
	}

	public void setF1(int f1) {
		this.f1 = f1;
	}

	public Integer getF2() {	
		return f2;
	}

	public void setF2(int f2) {
		this.f2 = f2;
	}

	public Integer getF3() {
		return f3;
	}

	public void setF3(int f3) {
		this.f3 = f3;
	}

	@Override
	public String toString() {
		return "Element" + element + ", l=" + l + ", a=" + a + ", h=" + h + ", sum=" + total + "]";
	}
}
