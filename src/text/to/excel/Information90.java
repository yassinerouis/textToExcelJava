package text.to.excel;

public class Information90 {
	String Matricule;
	int dossPaie;	
	String rubrique;	
	String PERIODE;	
	String NUMORD;
	public Information90(String matricule, int dossPaie, String rubrique, String pERIODE, String nUMORD) {
		super();
		Matricule = matricule;
		this.dossPaie = dossPaie;
		this.rubrique = rubrique;
		PERIODE = pERIODE;
		NUMORD = nUMORD;
	}
	public Information90() {
		super();
		// TODO Auto-generated constructor stub
	}
	public String getMatricule() {
		return Matricule;
	}
	public void setMatricule(String matricule) {
		Matricule = matricule;
	}
	public int getDossPaie() {
		return dossPaie;
	}
	public void setDossPaie(int dossPaie) {
		this.dossPaie = dossPaie;
	}
	public String getRubrique() {
		return rubrique;
	}
	public void setRubrique(String rubrique) {
		this.rubrique = rubrique;
	}
	public String getPERIODE() {
		return PERIODE;
	}
	public void setPERIODE(String pERIODE) {
		PERIODE = pERIODE;
	}
	public String getNUMORD() {
		return NUMORD;
	}
	public void setNUMORD(String nUMORD) {
		NUMORD = nUMORD;
	}
}
