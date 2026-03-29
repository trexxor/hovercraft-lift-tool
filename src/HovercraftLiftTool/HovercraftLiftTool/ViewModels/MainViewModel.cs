using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using HovercraftLiftTool.Models;
using HovercraftLiftTool.Services;

namespace HovercraftLiftTool.ViewModels;

// ViewModel for MainWindow.
// All displayed values are in whatever unit system is currently selected (US or SI).
// Internally, all inputs are converted to US Customary before calling HovercraftCalc.
public class MainViewModel : INotifyPropertyChanged
{
    // ---- Guards ----
    // Prevents recalculation from firing while we are programmatically writing results
    // back to the UI — same pattern as frmHovercraft.m_calculating in the VBA.
    private bool _calculating;

    // ---- Unit / pressure mode ----
    private bool _isUS = true;       // true = US Customary, false = SI
    private bool _useGauge = false;  // false = Absolute, true = Gauge

    // ---- Stored US and SI dictionaries ----
    // Populated after every calculation. Switching units reads from these
    // instead of recalculating, matching the VBA DisplayFromDictionary behavior.
    private Dictionary<string, double> _usValues  = new();
    private Dictionary<string, double> _siValues  = new();

    // ---- Displayed field values (in current unit system) ----
    // Each property raises PropertyChanged so the UI updates automatically.
    private double _pa, _rho;
    private double _wls, _wdw, _wtot;
    private double _lcg, _tcg, _vcgf, _vcgc;
    private double _lc, _bc, _ac;
    private double _dsk, _dag, _dc;
    private double _cd;
    private double _pca, _deltap;
    private double _perag, _aag;
    private double _vout, _vin, _vdot, _mdot;
    private double _eta, _pf;

    // ---- Unit label strings (update when unit system changes) ----
    private string _pressureUnit = "psia";
    private string _deltapUnit   = "psi";
    private string _densityUnit  = "slug/ft³";
    private string _weightUnit   = "LT";
    private string _lengthUnit   = "ft";
    private string _areaUnit     = "ft²";
    private string _velUnit      = "ft/s";
    private string _volFlowUnit  = "ft³/s";
    private string _massFlowUnit = "slug/s";
    private string _powerUnit    = "hp";

    // ---- Solve-for flags ----
    private bool _solveWls, _solveWdw, _solveWtot;
    private bool _solveLc, _solveBc, _solveAc;
    private bool _solveDsk, _solveDag, _solveDc;
    private bool _solvePca, _solveDeltaP;
    private bool _solveEta, _solvePf;

    // ---- Scenario table ----
    private int _scenarioCount = 0;
    public ObservableCollection<ScenarioRow> ScenarioLog { get; } = new();

    // ---- Commands ----
    public ICommand RecalculateCommand  { get; }
    public ICommand RecordScenarioCommand { get; }
    public ICommand ResetCommand        { get; }
    public ICommand ClearCommand        { get; }
    public ICommand ExportCsvCommand    { get; }

    public MainViewModel()
    {
        RecalculateCommand    = new RelayCommand(_ => RunCalculations());
        RecordScenarioCommand = new RelayCommand(_ => RecordScenario());
        ResetCommand          = new RelayCommand(_ => ResetForm());
        ClearCommand          = new RelayCommand(_ => ClearForm());
        ExportCsvCommand      = new RelayCommand(_ => ScenarioCsvExporter.Export(ScenarioLog));

        ApplyDefaults();
        RunCalculations();
    }

    // ============================================================
    // Unit system & pressure mode
    // ============================================================

    public bool IsUS
    {
        get => _isUS;
        set
        {
            if (_isUS == value) return;
            _isUS = value;
            OnPropertyChanged();
            // Toggle SI does NOT recalculate — just swaps displayed values
            // from the already-populated dictionaries (matches VBA behavior).
            if (!_calculating)
                DisplayFromDictionaries();
        }
    }

    public bool IsSI
    {
        get => !_isUS;
        set => IsUS = !value;
    }

    public bool IsAbsolute
    {
        get => !_useGauge;
        set
        {
            if (!_useGauge == value) return;  // already in requested state
            _useGauge = !value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsGauge));
            if (!_calculating) RunCalculations();
        }
    }

    public bool IsGauge
    {
        get => _useGauge;
        set
        {
            if (_useGauge == value) return;
            _useGauge = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsAbsolute));
            if (!_calculating) RunCalculations();
        }
    }

    // ============================================================
    // Input field properties — each setter triggers recalculation
    // on field exit (called from the View's LostFocus handler via
    // the RecalculateCommand binding, not from the setter directly).
    // The setter just stores the value and raises PropertyChanged.
    // ============================================================

    public double Pa    { get => _pa;   set { _pa   = value; OnPropertyChanged(); } }
    public double Rho   { get => _rho;  set { _rho  = value; OnPropertyChanged(); } }

    public double Wls   { get => _wls;  set { _wls  = value; OnPropertyChanged(); } }
    public double Wdw   { get => _wdw;  set { _wdw  = value; OnPropertyChanged(); } }
    public double Wtot  { get => _wtot; set { _wtot = value; OnPropertyChanged(); } }

    public double LCG   { get => _lcg;  set { _lcg  = value; OnPropertyChanged(); } }
    public double TCG   { get => _tcg;  set { _tcg  = value; OnPropertyChanged(); } }
    public double VCGf  { get => _vcgf; set { _vcgf = value; OnPropertyChanged(); } }
    public double VCGc  { get => _vcgc; set { _vcgc = value; OnPropertyChanged(); } }

    public double Lc    { get => _lc;   set { _lc   = value; OnPropertyChanged(); } }
    public double Bc    { get => _bc;   set { _bc   = value; OnPropertyChanged(); } }
    public double Ac    { get => _ac;   set { _ac   = value; OnPropertyChanged(); } }

    public double Dsk   { get => _dsk;  set { _dsk  = value; OnPropertyChanged(); } }
    public double Dag   { get => _dag;  set { _dag  = value; OnPropertyChanged(); } }
    public double Dc    { get => _dc;   set { _dc   = value; OnPropertyChanged(); } }

    public double Cd    { get => _cd;   set { _cd   = value; OnPropertyChanged(); } }

    public double Pca    { get => _pca;    set { _pca    = value; OnPropertyChanged(); } }
    public double DeltaP { get => _deltap; set { _deltap = value; OnPropertyChanged(); } }

    // Always-calculated (read-only in UI)
    public double Perag  { get => _perag; private set { _perag = value; OnPropertyChanged(); } }
    public double Aag    { get => _aag;   private set { _aag   = value; OnPropertyChanged(); } }
    public double Vin    { get => _vin;   private set { _vin   = value; OnPropertyChanged(); } }
    public double Vout   { get => _vout;  private set { _vout  = value; OnPropertyChanged(); } }
    public double Vdot   { get => _vdot;  private set { _vdot  = value; OnPropertyChanged(); } }
    public double Mdot   { get => _mdot;  private set { _mdot  = value; OnPropertyChanged(); } }

    public double Eta    { get => _eta;  set { _eta  = value; OnPropertyChanged(); } }
    public double Pf     { get => _pf;   set { _pf   = value; OnPropertyChanged(); } }

    // ============================================================
    // Solve-for flags — each setter triggers immediate recalculation
    // (matches VBA optSolveXxx_Click → RunCalculations).
    // ============================================================

    public bool SolveWls  { get => _solveWls;  set { _solveWls  = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }
    public bool SolveWdw  { get => _solveWdw;  set { _solveWdw  = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }
    public bool SolveWtot { get => _solveWtot; set { _solveWtot = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }

    public bool SolveLc   { get => _solveLc;   set { _solveLc   = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }
    public bool SolveBc   { get => _solveBc;   set { _solveBc   = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }
    public bool SolveAc   { get => _solveAc;   set { _solveAc   = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }

    public bool SolveDsk  { get => _solveDsk;  set { _solveDsk  = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }
    public bool SolveDag  { get => _solveDag;  set { _solveDag  = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }
    public bool SolveDc   { get => _solveDc;   set { _solveDc   = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }

    public bool SolvePca    { get => _solvePca;    set { _solvePca    = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }
    public bool SolveDeltaP { get => _solveDeltaP; set { _solveDeltaP = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }

    public bool SolveEta  { get => _solveEta;  set { _solveEta  = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }
    public bool SolvePf   { get => _solvePf;   set { _solvePf   = value; OnPropertyChanged(); if (!_calculating) RunCalculations(); } }

    // IsEnabled inverses for solved/reserved fields
    public bool WlsEnabled  => !_solveWls;
    public bool WdwEnabled  => !_solveWdw;
    public bool WtotEnabled => !_solveWtot;
    public bool LcEnabled   => !_solveLc;
    public bool BcEnabled   => !_solveBc;
    public bool AcEnabled   => !_solveAc;
    public bool DskEnabled  => !_solveDsk;
    public bool DagEnabled  => !_solveDag;
    public bool DcEnabled   => !_solveDc;
    public bool PcaEnabled    => !_solvePca;
    public bool DeltaPEnabled => !_solveDeltaP;
    public bool EtaEnabled  => !_solveEta;
    public bool PfEnabled   => !_solvePf;

    // ============================================================
    // Unit label strings
    // ============================================================

    public string PressureUnit  { get => _pressureUnit;  private set { _pressureUnit  = value; OnPropertyChanged(); } }
    public string DeltaPUnit    { get => _deltapUnit;    private set { _deltapUnit    = value; OnPropertyChanged(); } }
    public string DensityUnit   { get => _densityUnit;   private set { _densityUnit   = value; OnPropertyChanged(); } }
    public string WeightUnit    { get => _weightUnit;    private set { _weightUnit    = value; OnPropertyChanged(); } }
    public string LengthUnit    { get => _lengthUnit;    private set { _lengthUnit    = value; OnPropertyChanged(); } }
    public string AreaUnit      { get => _areaUnit;      private set { _areaUnit      = value; OnPropertyChanged(); } }
    public string VelUnit       { get => _velUnit;       private set { _velUnit       = value; OnPropertyChanged(); } }
    public string VolFlowUnit   { get => _volFlowUnit;   private set { _volFlowUnit   = value; OnPropertyChanged(); } }
    public string MassFlowUnit  { get => _massFlowUnit;  private set { _massFlowUnit  = value; OnPropertyChanged(); } }
    public string PowerUnit     { get => _powerUnit;     private set { _powerUnit     = value; OnPropertyChanged(); } }

    // ============================================================
    // RunCalculations — port of VBA RunCalculations()
    // ============================================================

    public void RunCalculations()
    {
        _calculating = true;
        try
        {
            // Convert displayed values back to US Customary base units for calculation.
            // If isUS, displayed values ARE US units. If SI, divide by conversion factors.
            var inputs = ReadInputsAsUS();
            var flags  = BuildSolveFlags();
            var out_   = HovercraftCalc.Calculate(inputs, flags, _useGauge);

            PopulateDictionaries(inputs, out_);
            DisplayFromDictionaries();
        }
        finally
        {
            _calculating = false;
        }
    }

    // Reads all displayed field values and converts to US Customary base units.
    private HovercraftInputs ReadInputsAsUS()
    {
        if (_isUS)
        {
            return new HovercraftInputs
            {
                P_a_psi  = _pa,
                Rho      = _rho,
                W_ls_LT  = _wls,
                W_dw_LT  = _wdw,
                W_tot_LT = _wtot,
                LCG      = _lcg,
                TCG      = _tcg,
                VCGf     = _vcgf,
                VCGc     = _vcgc,
                L_c      = _lc,
                B_c      = _bc,
                A_c      = _ac,
                D_sk     = _dsk,
                D_ag     = _dag,
                D_c      = _dc,
                C_d      = _cd,
                Eta      = _eta,
                P_f_hp   = _pf,
            };
        }
        else
        {
            // Convert from SI display values back to US Customary
            return new HovercraftInputs
            {
                P_a_psi  = UnitConverter.ToUS(_pa,  "pressure"),
                Rho      = UnitConverter.ToUS(_rho,  "density"),
                W_ls_LT  = UnitConverter.ToUS(_wls,  "weight"),
                W_dw_LT  = UnitConverter.ToUS(_wdw,  "weight"),
                W_tot_LT = UnitConverter.ToUS(_wtot, "weight"),
                LCG      = UnitConverter.ToUS(_lcg,  "length"),
                TCG      = UnitConverter.ToUS(_tcg,  "length"),
                VCGf     = UnitConverter.ToUS(_vcgf, "length"),
                VCGc     = UnitConverter.ToUS(_vcgc, "length"),
                L_c      = UnitConverter.ToUS(_lc,   "length"),
                B_c      = UnitConverter.ToUS(_bc,   "length"),
                A_c      = UnitConverter.ToUS(_ac,   "area"),
                D_sk     = UnitConverter.ToUS(_dsk,  "length"),
                D_ag     = UnitConverter.ToUS(_dag,  "length"),
                D_c      = UnitConverter.ToUS(_dc,   "length"),
                C_d      = _cd,   // dimensionless
                Eta      = _eta,  // dimensionless
                P_f_hp   = UnitConverter.ToUS(_pf,   "power"),
            };
        }
    }

    private SolveFlags BuildSolveFlags() => new()
    {
        SolveWls    = _solveWls,
        SolveWdw    = _solveWdw,
        SolveWtot   = _solveWtot,
        SolveLc     = _solveLc,
        SolveBc     = _solveBc,
        SolveAc     = _solveAc,
        SolveDsk    = _solveDsk,
        SolveDag    = _solveDag,
        SolveDc     = _solveDc,
        SolvePca    = _solvePca,
        SolveDeltaP = _solveDeltaP,
        SolveEta    = _solveEta,
        SolvePf     = _solvePf,
    };

    // ============================================================
    // PopulateDictionaries — mirrors VBA PopulateDictionaries()
    // Stores US and SI values for both input-that-were-solved and
    // computed outputs, so unit toggle requires no recalculation.
    // ============================================================

    private void PopulateDictionaries(HovercraftInputs inp, HovercraftOutputs o)
    {
        // US values (stored as US Customary display units)
        _usValues = new Dictionary<string, double>
        {
            ["Pa"]     = inp.P_a_psi,
            ["Rho"]    = inp.Rho,
            ["Wls"]    = o.W_ls_LT,
            ["Wdw"]    = o.W_dw_LT,
            ["Wtot"]   = o.W_tot_LT,
            ["LCG"]    = inp.LCG,
            ["TCG"]    = inp.TCG,
            ["VCGf"]   = inp.VCGf,
            ["VCGc"]   = inp.VCGc,
            ["Lc"]     = o.L_c,
            ["Bc"]     = o.B_c,
            ["Ac"]     = o.A_c,
            ["Dsk"]    = o.D_sk,
            ["Dag"]    = o.D_ag,
            ["Dc"]     = o.D_c,
            ["Cd"]     = inp.C_d,
            ["Pca"]    = o.P_ca_psi,
            ["DeltaP"] = o.Delta_p_psi,
            ["Perag"]  = o.Per_ag,
            ["Aag"]    = o.A_ag,
            ["Vin"]    = o.V_in,
            ["Vout"]   = o.V_out,
            ["Vdot"]   = o.V_dot,
            ["Mdot"]   = o.M_dot,
            ["Eta"]    = o.Eta,
            ["Pf"]     = o.P_f_hp,
        };

        // SI values (converted from US)
        _siValues = new Dictionary<string, double>
        {
            ["Pa"]     = UnitConverter.ToSI(inp.P_a_psi,  "pressure"),
            ["Rho"]    = UnitConverter.ToSI(inp.Rho,      "density"),
            ["Wls"]    = UnitConverter.ToSI(o.W_ls_LT,   "weight"),
            ["Wdw"]    = UnitConverter.ToSI(o.W_dw_LT,   "weight"),
            ["Wtot"]   = UnitConverter.ToSI(o.W_tot_LT,  "weight"),
            ["LCG"]    = UnitConverter.ToSI(inp.LCG,     "length"),
            ["TCG"]    = UnitConverter.ToSI(inp.TCG,     "length"),
            ["VCGf"]   = UnitConverter.ToSI(inp.VCGf,   "length"),
            ["VCGc"]   = UnitConverter.ToSI(inp.VCGc,   "length"),
            ["Lc"]     = UnitConverter.ToSI(o.L_c,      "length"),
            ["Bc"]     = UnitConverter.ToSI(o.B_c,      "length"),
            ["Ac"]     = UnitConverter.ToSI(o.A_c,      "area"),
            ["Dsk"]    = UnitConverter.ToSI(o.D_sk,     "length"),
            ["Dag"]    = UnitConverter.ToSI(o.D_ag,     "length"),
            ["Dc"]     = UnitConverter.ToSI(o.D_c,      "length"),
            ["Cd"]     = inp.C_d,
            ["Pca"]    = UnitConverter.ToSI(o.P_ca_psi,   "pressure"),
            ["DeltaP"] = UnitConverter.ToSI(o.Delta_p_psi,"pressure"),
            ["Perag"]  = UnitConverter.ToSI(o.Per_ag,   "length"),
            ["Aag"]    = UnitConverter.ToSI(o.A_ag,     "area"),
            ["Vin"]    = UnitConverter.ToSI(o.V_in,     "velocity"),
            ["Vout"]   = UnitConverter.ToSI(o.V_out,    "velocity"),
            ["Vdot"]   = UnitConverter.ToSI(o.V_dot,    "volflow"),
            ["Mdot"]   = UnitConverter.ToSI(o.M_dot,    "massflow"),
            ["Eta"]    = o.Eta,
            ["Pf"]     = UnitConverter.ToSI(o.P_f_hp,  "power"),
        };
    }

    // ============================================================
    // DisplayFromDictionaries — mirrors VBA DisplayFromDictionary()
    // Reads from US or SI dictionary and updates all bound properties.
    // Also updates unit label strings.
    // ============================================================

    private void DisplayFromDictionaries()
    {
        var vals = _isUS ? _usValues : _siValues;
        if (vals.Count == 0) return;

        // Update unit labels
        PressureUnit  = UnitConverter.PressureUnit(_isUS);
        DeltaPUnit    = UnitConverter.DeltaPUnit(_isUS);
        DensityUnit   = UnitConverter.DensityUnit(_isUS);
        WeightUnit    = UnitConverter.WeightUnit(_isUS);
        LengthUnit    = UnitConverter.LengthUnit(_isUS);
        AreaUnit      = UnitConverter.AreaUnit(_isUS);
        VelUnit       = UnitConverter.VelocityUnit(_isUS);
        VolFlowUnit   = UnitConverter.VolFlowUnit(_isUS);
        MassFlowUnit  = UnitConverter.MassFlowUnit(_isUS);
        PowerUnit     = UnitConverter.PowerUnit(_isUS);

        // Write values back to display properties
        _pa     = vals["Pa"];
        _rho    = vals["Rho"];
        _wls    = vals["Wls"];
        _wdw    = vals["Wdw"];
        _wtot   = vals["Wtot"];
        _lcg    = vals["LCG"];
        _tcg    = vals["TCG"];
        _vcgf   = vals["VCGf"];
        _vcgc   = vals["VCGc"];
        _lc     = vals["Lc"];
        _bc     = vals["Bc"];
        _ac     = vals["Ac"];
        _dsk    = vals["Dsk"];
        _dag    = vals["Dag"];
        _dc     = vals["Dc"];
        _cd     = vals["Cd"];
        _pca    = vals["Pca"];
        _deltap = vals["DeltaP"];
        _perag  = vals["Perag"];
        _aag    = vals["Aag"];
        _vin    = vals["Vin"];
        _vout   = vals["Vout"];
        _vdot   = vals["Vdot"];
        _mdot   = vals["Mdot"];
        _eta    = vals["Eta"];
        _pf     = vals["Pf"];

        // Raise PropertyChanged for all displayed fields at once
        OnPropertyChanged(nameof(Pa));
        OnPropertyChanged(nameof(Rho));
        OnPropertyChanged(nameof(Wls));
        OnPropertyChanged(nameof(Wdw));
        OnPropertyChanged(nameof(Wtot));
        OnPropertyChanged(nameof(LCG));
        OnPropertyChanged(nameof(TCG));
        OnPropertyChanged(nameof(VCGf));
        OnPropertyChanged(nameof(VCGc));
        OnPropertyChanged(nameof(Lc));
        OnPropertyChanged(nameof(Bc));
        OnPropertyChanged(nameof(Ac));
        OnPropertyChanged(nameof(Dsk));
        OnPropertyChanged(nameof(Dag));
        OnPropertyChanged(nameof(Dc));
        OnPropertyChanged(nameof(Cd));
        OnPropertyChanged(nameof(Pca));
        OnPropertyChanged(nameof(DeltaP));
        OnPropertyChanged(nameof(Perag));
        OnPropertyChanged(nameof(Aag));
        OnPropertyChanged(nameof(Vin));
        OnPropertyChanged(nameof(Vout));
        OnPropertyChanged(nameof(Vdot));
        OnPropertyChanged(nameof(Mdot));
        OnPropertyChanged(nameof(Eta));
        OnPropertyChanged(nameof(Pf));

        // Refresh enabled states
        OnPropertyChanged(nameof(WlsEnabled));
        OnPropertyChanged(nameof(WdwEnabled));
        OnPropertyChanged(nameof(WtotEnabled));
        OnPropertyChanged(nameof(LcEnabled));
        OnPropertyChanged(nameof(BcEnabled));
        OnPropertyChanged(nameof(AcEnabled));
        OnPropertyChanged(nameof(DskEnabled));
        OnPropertyChanged(nameof(DagEnabled));
        OnPropertyChanged(nameof(DcEnabled));
        OnPropertyChanged(nameof(PcaEnabled));
        OnPropertyChanged(nameof(DeltaPEnabled));
        OnPropertyChanged(nameof(EtaEnabled));
        OnPropertyChanged(nameof(PfEnabled));
    }

    // ============================================================
    // RecordScenario — mirrors VBA RecordScenario()
    // ============================================================

    private void RecordScenario()
    {
        _scenarioCount++;
        var row = new ScenarioRow
        {
            ScenarioNum  = _scenarioCount,
            Timestamp    = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            Description  = "",
            UnitSystem   = _isUS ? "US Customary" : "SI",
            PressureMode = _useGauge ? "Gauge" : "Absolute",
            W_ls    = _wls,
            W_dw    = _wdw,
            W_tot   = _wtot,
            LCG     = _lcg,
            TCG     = _tcg,
            VCGf    = _vcgf,
            VCGc    = _vcgc,
            P_a     = _pa,
            Rho     = _rho,
            L_c     = _lc,
            B_c     = _bc,
            A_c     = _ac,
            D_sk    = _dsk,
            D_ag    = _dag,
            D_c     = _dc,
            Per_ag  = _perag,
            A_ag    = _aag,
            P_ca    = _pca,
            Delta_p = _deltap,
            V_in    = _vin,
            V_out   = _vout,
            V_dot   = _vdot,
            M_dot   = _mdot,
            C_d     = _cd,
            Eta     = _eta,
            P_f     = _pf,
        };
        ScenarioLog.Add(row);

        MessageBox.Show(
            $"Scenario {_scenarioCount} recorded.\nYou can view and edit the description in the scenario table below.",
            "Scenario Recorded", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    // ============================================================
    // ResetForm / ClearForm — mirrors VBA ResetForm() / ClearForm()
    // ============================================================

    private void ResetForm()
    {
        if (MessageBox.Show("Reset all fields to default values?",
            "Reset", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
            return;

        ApplyDefaults();
        RunCalculations();
    }

    private void ClearForm()
    {
        if (MessageBox.Show("Clear all input values?",
            "Clear", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
            return;

        _calculating = true;
        _wls  = 0; _wdw  = 0; _wtot = 0;
        _lcg  = 0; _tcg  = 0; _vcgf = 0; _vcgc = 0;
        _lc   = 0; _bc   = 0; _ac   = 0;
        _dsk  = 0; _dag  = 0; _dc   = 0;
        _cd   = 1.0;
        _eta  = 1.0;
        _pf   = 0;
        _calculating = false;

        RunCalculations();
    }

    // Set all default values matching VBA ResetForm() and UserForm_Initialize()
    private void ApplyDefaults()
    {
        _calculating = true;

        _isUS     = true;
        _useGauge = false;
        OnPropertyChanged(nameof(IsUS));
        OnPropertyChanged(nameof(IsSI));
        OnPropertyChanged(nameof(IsAbsolute));
        OnPropertyChanged(nameof(IsGauge));

        _pa   = 14.696;
        _rho  = 0.00237;
        _wls  = 0.0;   _wdw = 0.0;   _wtot = 0.0;
        _lcg  = 0.0;   _tcg = 0.0;   _vcgf = 0.0;  _vcgc = 0.0;
        _lc   = 0.0;   _bc  = 0.0;   _ac   = 0.0;
        _dsk  = 0.0;   _dag = 0.0;   _dc   = 0.0;
        _cd   = 1.0;
        _eta  = 1.0;
        _pf   = 0.0;

        // Default solve-for selections (matching VBA UserForm_Initialize defaults)
        _solveWls  = false; _solveWdw  = false; _solveWtot = true;
        _solveLc   = false; _solveBc   = false; _solveAc   = true;
        _solveDsk  = false; _solveDag  = false; _solveDc   = true;
        _solvePca  = false; _solveDeltaP = true;
        _solveEta  = false; _solvePf  = true;

        OnPropertyChanged(nameof(SolveWls));  OnPropertyChanged(nameof(SolveWdw));  OnPropertyChanged(nameof(SolveWtot));
        OnPropertyChanged(nameof(SolveLc));   OnPropertyChanged(nameof(SolveBc));   OnPropertyChanged(nameof(SolveAc));
        OnPropertyChanged(nameof(SolveDsk));  OnPropertyChanged(nameof(SolveDag));  OnPropertyChanged(nameof(SolveDc));
        OnPropertyChanged(nameof(SolvePca));  OnPropertyChanged(nameof(SolveDeltaP));
        OnPropertyChanged(nameof(SolveEta));  OnPropertyChanged(nameof(SolvePf));

        _calculating = false;
    }

    // ============================================================
    // INotifyPropertyChanged
    // ============================================================

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

// Minimal ICommand implementation — no external dependencies needed.
internal class RelayCommand : ICommand
{
    private readonly Action<object?> _execute;
    private readonly Func<object?, bool>? _canExecute;

    public RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null)
    {
        _execute    = execute;
        _canExecute = canExecute;
    }

    public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;
    public void Execute(object? parameter)    => _execute(parameter);
    public event EventHandler? CanExecuteChanged
    {
        add    => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }
}
