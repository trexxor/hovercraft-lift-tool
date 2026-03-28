namespace HovercraftLiftTool.Models;

// All outputs in US Customary base units (same convention as HovercraftInputs).
// The ViewModel converts to display units using UnitConverter before showing them.
public class HovercraftOutputs
{
    // Inputs that were back-solved (written back to inputs after solve)
    public double W_ls_LT  { get; set; }
    public double W_dw_LT  { get; set; }
    public double W_tot_LT { get; set; }
    public double L_c      { get; set; }
    public double B_c      { get; set; }
    public double A_c      { get; set; }
    public double D_sk     { get; set; }
    public double D_ag     { get; set; }
    public double D_c      { get; set; }
    public double P_ca_psi { get; set; }   // psia display
    public double Delta_p_psi { get; set; } // psi display
    public double Eta      { get; set; }
    public double P_f_hp   { get; set; }

    // Derived geometry (always calculated)
    public double Per_ag { get; set; }  // ft
    public double A_ag   { get; set; }  // ft²

    // Flow results (always calculated)
    public double V_out { get; set; }   // ft/s (Torricelli)
    public double V_in  { get; set; }   // ft/s (reserved = 0)
    public double V_dot { get; set; }   // ft³/s
    public double M_dot { get; set; }   // slug/s
}
