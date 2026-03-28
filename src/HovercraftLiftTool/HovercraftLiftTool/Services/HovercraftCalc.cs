using HovercraftLiftTool.Models;

namespace HovercraftLiftTool.Services;

// Solve-for flags — one per variable that supports solve-for.
// Multiple flags can be true simultaneously (one per equation group).
public class SolveFlags
{
    // Weight group (W_tot = W_ls + W_dw)
    public bool SolveWls  { get; set; }
    public bool SolveWdw  { get; set; }
    public bool SolveWtot { get; set; }

    // Cushion area group (A_c = L_c * B_c)
    public bool SolveLc { get; set; }
    public bool SolveBc { get; set; }
    public bool SolveAc { get; set; }

    // Depth group (D_c = D_sk + D_ag)
    public bool SolveDsk { get; set; }
    public bool SolveDag { get; set; }
    public bool SolveDc  { get; set; }

    // Pressure group
    public bool SolvePca    { get; set; }
    public bool SolveDeltaP { get; set; }

    // Fan power group (P_f = delta_p * V_dot / eta)
    public bool SolveEta { get; set; }
    public bool SolvePf  { get; set; }
}

// Direct C# port of HovercraftCalc.bas RunCalculations().
// All inputs and outputs are in US Customary base units:
//   Weight    : lb  (W_*_LT fields are Long Tons; multiply by 2240 internally)
//   Pressure  : lb/ft²  (P_a_psi in psia; multiply by 144 for lb/ft²)
//   Length    : ft
//   Area      : ft²
//   Density   : slug/ft³
//   Power     : ft·lb/s internally (displayed as hp = /550)
public static class HovercraftCalc
{
    private const double LB_PER_LT   = 2240.0;
    private const double IN2_PER_FT2 = 144.0;
    private const double FTLBS_PER_HP = 550.0;

    // Runs the full calculation sequence and returns outputs.
    // inputs is not modified — outputs contains all solved/derived values.
    public static HovercraftOutputs Calculate(HovercraftInputs inputs, SolveFlags solve, bool useGauge)
    {
        // Working copies of all variables in US Customary base units
        double p_a   = inputs.P_a_psi * IN2_PER_FT2;   // lb/ft²
        double rho   = inputs.Rho;                      // slug/ft³

        double W_ls_LT  = inputs.W_ls_LT;
        double W_dw_LT  = inputs.W_dw_LT;
        double W_tot_LT = inputs.W_tot_LT;

        double W_ls  = W_ls_LT  * LB_PER_LT;           // lb
        double W_dw  = W_dw_LT  * LB_PER_LT;
        double W_tot = W_tot_LT * LB_PER_LT;

        double L_c  = inputs.L_c;
        double B_c  = inputs.B_c;
        double A_c  = inputs.A_c;

        double D_sk = inputs.D_sk;
        double D_ag = inputs.D_ag;
        double D_c  = inputs.D_c;

        double C_d  = inputs.C_d;
        double eta  = inputs.Eta;
        double P_f_hp = inputs.P_f_hp;

        // ---- Weight group: W_tot = W_ls + W_dw ----
        if (solve.SolveWtot)
        {
            W_tot    = W_ls + W_dw;
            W_tot_LT = W_tot / LB_PER_LT;
        }
        else if (solve.SolveWls)
        {
            W_ls    = W_tot - W_dw;
            W_ls_LT = W_ls / LB_PER_LT;
        }
        else if (solve.SolveWdw)
        {
            W_dw    = W_tot - W_ls;
            W_dw_LT = W_dw / LB_PER_LT;
        }

        // ---- Cushion area group: A_c = L_c * B_c ----
        if (solve.SolveAc)
        {
            A_c = L_c * B_c;
        }
        else if (solve.SolveLc)
        {
            if (B_c != 0) L_c = A_c / B_c;
        }
        else if (solve.SolveBc)
        {
            if (L_c != 0) B_c = A_c / L_c;
        }

        // ---- Depth group: D_c = D_sk + D_ag ----
        if (solve.SolveDc)
        {
            D_c = D_sk + D_ag;
        }
        else if (solve.SolveDsk)
        {
            D_sk = D_c - D_ag;
        }
        else if (solve.SolveDag)
        {
            D_ag = D_c - D_sk;
        }

        // ---- Derived geometry (always calculated) ----
        double Per_ag = 2.0 * (L_c + B_c);   // ft
        double A_ag   = Per_ag * D_ag;         // ft²

        // ---- Pressure group ----
        // Gauge mode:    p_ca = W_tot / A_c  (gauge)
        //                delta_p = p_ca
        // Absolute mode: p_ca = (W_tot / A_c) + p_a  (absolute)
        //                delta_p = p_ca - p_a
        double p_ca   = 0.0;   // lb/ft²
        double delta_p = 0.0;  // lb/ft²

        if (A_c != 0)
        {
            if (useGauge)
            {
                p_ca    = W_tot / A_c;
                delta_p = p_ca;
            }
            else
            {
                p_ca    = (W_tot / A_c) + p_a;
                delta_p = p_ca - p_a;
            }
        }

        // ---- Flow (always calculated) ----
        // Torricelli jet velocity: v_out = C_d * sqrt(2 * delta_p / rho)
        double v_out = 0.0;
        if (rho > 0 && delta_p >= 0)
            v_out = C_d * Math.Sqrt(2.0 * delta_p / rho);

        double V_dot = A_ag * v_out;    // ft³/s
        double m_dot = V_dot * rho;     // slug/s
        double v_in  = 0.0;             // reserved

        // ---- Fan power group: P_f = (delta_p * V_dot) / eta ----
        if (solve.SolvePf)
        {
            if (eta > 0)
            {
                double P_f = (delta_p * V_dot) / eta;  // ft·lb/s
                P_f_hp = P_f / FTLBS_PER_HP;
            }
            else
            {
                P_f_hp = 0.0;
            }
        }
        else if (solve.SolveEta)
        {
            double P_f = P_f_hp * FTLBS_PER_HP;
            eta = (P_f > 0) ? (delta_p * V_dot) / P_f : 0.0;
        }

        // Convert lb/ft² pressures back to psia / psi for display
        double p_ca_psi    = p_ca    / IN2_PER_FT2;
        double delta_p_psi = delta_p / IN2_PER_FT2;

        return new HovercraftOutputs
        {
            W_ls_LT    = W_ls_LT,
            W_dw_LT    = W_dw_LT,
            W_tot_LT   = W_tot_LT,
            L_c        = L_c,
            B_c        = B_c,
            A_c        = A_c,
            D_sk       = D_sk,
            D_ag       = D_ag,
            D_c        = D_c,
            P_ca_psi   = p_ca_psi,
            Delta_p_psi = delta_p_psi,
            Eta        = eta,
            P_f_hp     = P_f_hp,
            Per_ag     = Per_ag,
            A_ag       = A_ag,
            V_out      = v_out,
            V_in       = v_in,
            V_dot      = V_dot,
            M_dot      = m_dot,
        };
    }
}
