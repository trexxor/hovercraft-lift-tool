namespace HovercraftLiftTool.Models;

// Matches the 31 columns written by RecordScenario in the VBA.
// Values are stored exactly as displayed (in whatever unit system was active when recorded).
public class ScenarioRow
{
    public int    ScenarioNum  { get; set; }
    public string Timestamp    { get; set; } = "";
    public string Description  { get; set; } = "";
    public string UnitSystem   { get; set; } = "";
    public string PressureMode { get; set; } = "";

    public double W_ls   { get; set; }
    public double W_dw   { get; set; }
    public double W_tot  { get; set; }
    public double LCG    { get; set; }
    public double TCG    { get; set; }
    public double VCGf   { get; set; }
    public double VCGc   { get; set; }

    public double P_a    { get; set; }
    public double Rho    { get; set; }

    public double L_c    { get; set; }
    public double B_c    { get; set; }
    public double A_c    { get; set; }
    public double D_sk   { get; set; }
    public double D_ag   { get; set; }
    public double D_c    { get; set; }
    public double Per_ag { get; set; }
    public double A_ag   { get; set; }

    public double P_ca    { get; set; }
    public double Delta_p { get; set; }
    public double V_in    { get; set; }
    public double V_out   { get; set; }
    public double V_dot   { get; set; }
    public double M_dot   { get; set; }
    public double C_d     { get; set; }

    public double Eta  { get; set; }
    public double P_f  { get; set; }
}
