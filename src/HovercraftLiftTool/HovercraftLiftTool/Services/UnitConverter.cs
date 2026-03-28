namespace HovercraftLiftTool.Services;

// Mirrors the ToSI / ToUS functions in VBA HovercraftCalc.bas.
// All conversion factors are identical to the VBA source (NIST SP811).
//
// US Customary base units used throughout:
//   Pressure  : psia (display) / lb/ft² (internal)
//   Weight    : Long Tons (display) / lb (internal)
//   Length    : ft
//   Area      : ft²
//   Density   : slug/ft³
//   Mass flow : slug/s
//   Vol flow  : ft³/s
//   Velocity  : ft/s
//   Power     : hp
public static class UnitConverter
{
    // ---- Conversion factors (US → SI, multiply) ----
    public const double FT_TO_M          = 0.3048;             // exact
    public const double FT2_TO_M2        = 0.09290304;         // exact, derived
    public const double FT3_TO_M3        = 0.028316847;        // derived
    public const double SLUG_TO_KG       = 14.593903;
    public const double SLUGFT3_TO_KGM3  = 515.3788;           // slug/ft³ → kg/m³
    public const double LT_TO_T          = 1.0160469088;       // long tons → metric tonnes
    public const double LBFT2_TO_PA      = 47.88026;           // lb/ft² → Pa (not used in display path)
    public const double PSI_TO_PA        = 6894.757;           // psi / psia → Pa
    public const double FPS_TO_MPS       = 0.3048;             // ft/s → m/s (same as FT_TO_M)
    public const double FT3S_TO_M3S      = 0.028316847;        // ft³/s → m³/s
    public const double HP_TO_W          = 745.699872;         // mechanical hp → W

    // Convert a value from US Customary display units to SI display units.
    // unitType strings match those used in the VBA dictionaries.
    public static double ToSI(double value, string unitType)
    {
        return unitType switch
        {
            "length"   => value * FT_TO_M,
            "area"     => value * FT2_TO_M2,
            "volume"   => value * FT3_TO_M3,
            "density"  => value * SLUGFT3_TO_KGM3,
            "mass"     => value * SLUG_TO_KG,
            "massflow" => value * SLUG_TO_KG,
            "weight"   => value * LT_TO_T,
            "pressure" => value * PSI_TO_PA,
            "velocity" => value * FPS_TO_MPS,
            "volflow"  => value * FT3S_TO_M3S,
            "power"    => value * HP_TO_W,
            _          => value   // dimensionless (eta, Cd) — no conversion
        };
    }

    // Convert a value from SI display units back to US Customary display units.
    public static double ToUS(double value, string unitType)
    {
        return unitType switch
        {
            "length"   => value / FT_TO_M,
            "area"     => value / FT2_TO_M2,
            "volume"   => value / FT3_TO_M3,
            "density"  => value / SLUGFT3_TO_KGM3,
            "mass"     => value / SLUG_TO_KG,
            "massflow" => value / SLUG_TO_KG,
            "weight"   => value / LT_TO_T,
            "pressure" => value / PSI_TO_PA,
            "velocity" => value / FPS_TO_MPS,
            "volflow"  => value / FT3S_TO_M3S,
            "power"    => value / HP_TO_W,
            _          => value
        };
    }

    // SI unit label strings for each field, keyed by field name.
    // Returns the US label when isUS=true, SI label when false.
    public static string PressureUnit(bool isUS)  => isUS ? "psia" : "Pa";
    public static string DeltaPUnit(bool isUS)    => isUS ? "psi"  : "Pa";
    public static string DensityUnit(bool isUS)   => isUS ? "slug/ft³" : "kg/m³";
    public static string WeightUnit(bool isUS)    => isUS ? "LT"   : "t";
    public static string LengthUnit(bool isUS)    => isUS ? "ft"   : "m";
    public static string AreaUnit(bool isUS)      => isUS ? "ft²"  : "m²";
    public static string VelocityUnit(bool isUS)  => isUS ? "ft/s" : "m/s";
    public static string VolFlowUnit(bool isUS)   => isUS ? "ft³/s" : "m³/s";
    public static string MassFlowUnit(bool isUS)  => isUS ? "slug/s" : "kg/s";
    public static string PowerUnit(bool isUS)     => isUS ? "hp"   : "W";
}
