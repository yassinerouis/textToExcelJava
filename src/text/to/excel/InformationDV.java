package text.to.excel;

public class InformationDV {
	String matricule;
	String date;
	public String getMatricule() {
		return matricule;
	}
	public void setMatricule(String matricule) {
		this.matricule = matricule;
	}
	public String getDate() {
		return date;
	}
	public void setDate(String date) {
		this.date = date;
	}
	public InformationDV(String matricule, String date) {
		super();
		this.matricule = matricule;
		this.date = date;
	}
	public InformationDV() {
		super();
		// TODO Auto-generated constructor stub
	}
}
