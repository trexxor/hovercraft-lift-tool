using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using HovercraftLiftTool.Models;
using Microsoft.Win32;

namespace HovercraftLiftTool.Services;

public static class ScenarioCsvExporter
{
    private static readonly string[] Headers =
    [
        "Scenario#", "Timestamp", "Description", "UnitSystem", "PressureMode",
        "W_ls", "W_dw", "W_tot", "LCG", "TCG", "VCGf", "VCGc",
        "p_a", "rho",
        "L_c", "B_c", "A_c", "D_sk", "D_ag", "D_c", "Per_ag", "A_ag",
        "p_ca", "delta_p", "v_in", "v_out", "V_dot", "m_dot", "C_d",
        "eta", "P_f"
    ];

    public static void Export(ObservableCollection<ScenarioRow> rows)
    {
        if (rows.Count == 0)
        {
            MessageBox.Show("No scenarios to export.", "Export CSV",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new SaveFileDialog
        {
            Filter   = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            FileName = "hovercraft_scenarios.csv",
            Title    = "Export Scenario Log"
        };

        if (dlg.ShowDialog() != true) return;

        var sb = new StringBuilder();
        sb.AppendLine(string.Join(",", Headers));

        foreach (var r in rows)
        {
            sb.AppendLine(string.Join(",",
                r.ScenarioNum,
                Quote(r.Timestamp),
                Quote(r.Description),
                Quote(r.UnitSystem),
                Quote(r.PressureMode),
                Fmt(r.W_ls),  Fmt(r.W_dw),  Fmt(r.W_tot),
                Fmt(r.LCG),   Fmt(r.TCG),   Fmt(r.VCGf),  Fmt(r.VCGc),
                Fmt(r.P_a),   Fmt(r.Rho),
                Fmt(r.L_c),   Fmt(r.B_c),   Fmt(r.A_c),
                Fmt(r.D_sk),  Fmt(r.D_ag),  Fmt(r.D_c),
                Fmt(r.Per_ag),Fmt(r.A_ag),
                Fmt(r.P_ca),  Fmt(r.Delta_p),
                Fmt(r.V_in),  Fmt(r.V_out),
                Fmt(r.V_dot), Fmt(r.M_dot), Fmt(r.C_d),
                Fmt(r.Eta),   Fmt(r.P_f)));
        }

        File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
        MessageBox.Show($"Exported {rows.Count} scenario(s) to:\n{dlg.FileName}",
            "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private static string Fmt(double v) => v.ToString("G6");
    private static string Quote(string s) => $"\"{s.Replace("\"", "\"\"")}\"";
}
