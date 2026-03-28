namespace HovercraftLiftTool.Models;

// All input values stored in US Customary base units:
//   Pressure  : lb/ft²  (p_a and p_ca internally; displayed as psia)
//   Weight    : lb      (displayed as Long Tons)
//   Length    : ft
//   Area      : ft²
//   Density   : slug/ft³
//   Flow      : ft³/s (volumetric), slug/s (mass)
//   Power     : hp
//   Velocity  : ft/s
public class HovercraftInputs
{
    // Atmospheric
    public double P_a_psi { get; set; } = 14.696;   // psia (display), lb/ft² (internal = P_a_psi * 144)
    public double Rho      { get; set; } = 0.00237;  // slug/ft³

    // Weight group (Long Tons for display; *2240 for lb internally)
    public double W_ls_LT  { get; set; } = 0.0;
    public double W_dw_LT  { get; set; } = 0.0;
    public double W_tot_LT { get; set; } = 0.0;

    // Center of gravity (ft; display only, not used in equations yet)
    public double LCG  { get; set; } = 0.0;
    public double TCG  { get; set; } = 0.0;
    public double VCGf { get; set; } = 0.0;
    public double VCGc { get; set; } = 0.0;

    // Cushion geometry (ft / ft²)
    public double L_c { get; set; } = 0.0;
    public double B_c { get; set; } = 0.0;
    public double A_c { get; set; } = 0.0;

    // Depth group (ft)
    public double D_sk { get; set; } = 0.0;
    public double D_ag { get; set; } = 0.0;
    public double D_c  { get; set; } = 0.0;

    // Flow coefficient (dimensionless)
    public double C_d { get; set; } = 1.0;

    // Fan power group
    public double Eta   { get; set; } = 1.0;    // dimensionless
    public double P_f_hp { get; set; } = 0.0;   // hp
}
