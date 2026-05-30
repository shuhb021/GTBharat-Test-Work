/**
 * FAR Engine — Ratio Calculator
 * Computes 8 financial ratios from Balance Sheet and P&L summary data.
 */

#include "far_engine.h"
#include "nlohmann/json.hpp"
#include <cmath>
#include <string>
#include <sstream>
#include <vector>
#include <algorithm>

using json = nlohmann::json;

static std::string to_lower(const std::string& s) {
    std::string result = s;
    std::transform(result.begin(), result.end(), result.begin(),
                   [](unsigned char c){ return std::tolower(c); });
    return result;
}

static void find_value(const json& rows, const std::string& keyword, double& cy, double& py, bool exact = false) {
    std::string kw = to_lower(keyword);
    for (const auto& row : rows) {
        std::string p = to_lower(row.value("particulars", ""));
        if (exact) {
            if (p == kw) {
                cy = row.value("cy", 0.0);
                py = row.value("py", 0.0);
                return;
            }
        } else {
            if (p.find(kw) != std::string::npos) {
                cy = row.value("cy", 0.0);
                py = row.value("py", 0.0);
                return;
            }
        }
    }
    cy = 0.0;
    py = 0.0;
}

static void find_liability_by_note(const json& rows, const std::string& note_num, bool is_current, double& cy_total, double& py_total) {
    cy_total = 0.0;
    py_total = 0.0;
    bool in_current_section = false;
    for (const auto& row : rows) {
        std::string p = to_lower(row.value("particulars", ""));
        if (p.find("current liabilities") != std::string::npos && p.find("non") == std::string::npos) {
            in_current_section = true;
        } else if (p.find("non current liabilities") != std::string::npos || p.find("non-current liabilities") != std::string::npos) {
            in_current_section = false;
        }
        
        std::string note = "";
        if (row.contains("note")) {
            if (row["note"].is_number()) {
                note = std::to_string(row["note"].get<int>());
            } else if (row["note"].is_string()) {
                note = row["note"].get<std::string>();
            }
        }
        
        if (note == note_num && in_current_section == is_current) {
            cy_total += row.value("cy", 0.0);
            py_total += row.value("py", 0.0);
        }
    }
}

static thread_local std::string g_ratio_result;

struct RatioResult {
    std::string key;
    std::string formula;
    double cy;
    double py;
    double change;
    std::string display_change;
    bool flag;
};

static RatioResult calc_ratio(
    const std::string& key,
    const std::string& formula,
    double cy_num, double cy_den,
    double py_num, double py_den,
    bool is_pct = false
) {
    RatioResult r;
    r.key = key;
    r.formula = formula;

    r.cy = (cy_den == 0.0) ? 0.0 : cy_num / cy_den;
    r.py = (py_den == 0.0) ? 0.0 : py_num / py_den;

    if (is_pct) {
        r.cy *= 100.0;
        r.py *= 100.0;
    }

    if (r.py == 0.0) {
        r.change = (r.cy == 0.0) ? 0.0 : 100.0;
        r.display_change = (r.cy == 0.0) ? "0.0%" : "N/A";
    } else {
        r.change = ((r.cy - r.py) / std::fabs(r.py)) * 100.0;
        std::ostringstream oss;
        oss.precision(1);
        oss << std::fixed << r.change << "%";
        r.display_change = oss.str();
    }

    r.flag = std::fabs(r.change) >= 10.0;

    // Round values
    r.cy = std::round(r.cy * 10000.0) / 10000.0;
    r.py = std::round(r.py * 10000.0) / 10000.0;
    r.change = std::round(r.change * 100.0) / 100.0;

    return r;
}

extern "C" {

FAR_EXPORT const char* compute_ratios(const char* bs_json, const char* pl_json) {
    try {
        json bs = json::parse(bs_json);
        json pl = json::parse(pl_json);

        // Extract BS values
        double total_equity_cy = bs.value("total_equity_cy", 0.0);
        double total_equity_py = bs.value("total_equity_py", 0.0);
        double ltb_cy = bs.value("ltb_cy", 0.0);
        double ltb_py = bs.value("ltb_py", 0.0);
        double stb_cy = bs.value("stb_cy", 0.0);
        double stb_py = bs.value("stb_py", 0.0);
        double total_ca_cy = bs.value("total_ca_cy", 0.0);
        double total_ca_py = bs.value("total_ca_py", 0.0);
        double total_cl_cy = bs.value("total_cl_cy", 0.0);
        double total_cl_py = bs.value("total_cl_py", 0.0);
        double total_ncl_cy = bs.value("total_ncl_cy", 0.0);
        double total_ncl_py = bs.value("total_ncl_py", 0.0);
        double debtors_cy = bs.value("debtors_cy", 0.0);
        double debtors_py = bs.value("debtors_py", 0.0);

        // Extract PL values
        double revenue_ops_cy = pl.value("revenue_ops_cy", 0.0);
        double revenue_ops_py = pl.value("revenue_ops_py", 0.0);
        double total_revenue_cy = pl.value("total_revenue_cy", 0.0);
        double total_revenue_py = pl.value("total_revenue_py", 0.0);
        double pbt_cy = pl.value("pbt_cy", 0.0);
        double pbt_py = pl.value("pbt_py", 0.0);
        double finance_cost_cy = pl.value("finance_cost_cy", 0.0);
        double finance_cost_py = pl.value("finance_cost_py", 0.0);

        std::vector<RatioResult> ratios = {
            calc_ratio("Debt Equity Ratio", "(LTB + STB) / Total Equity",
                       ltb_cy + stb_cy, total_equity_cy,
                       ltb_py + stb_py, total_equity_py),

            calc_ratio("Current Ratio", "Total Assets / Total Liabilities",
                       total_ca_cy, total_cl_cy,
                       total_ca_py, total_cl_py),

            calc_ratio("GCA Days", "(Total CA / Total Revenue) * 365",
                       total_ca_cy * 365.0, total_revenue_cy,
                       total_ca_py * 365.0, total_revenue_py),

            calc_ratio("Total Outside Liab. vs Total Equity", "(TCL + TNCL) / Total Equity",
                       total_cl_cy + total_ncl_cy, total_equity_cy,
                       total_cl_py + total_ncl_py, total_equity_py),

            calc_ratio("Net Profit Ratio", "PBT / Revenue from Ops",
                       pbt_cy, revenue_ops_cy,
                       pbt_py, revenue_ops_py, true),

            calc_ratio("Return on Equity (ROE)", "PBT / Avg. Total Equity",
                       pbt_cy, (total_equity_cy + total_equity_py) / 2.0,
                       pbt_py, total_equity_py / 2.0, true),

            calc_ratio("Return on Capital Employed (ROCE)", "(PBT + Finance Cost) / (TE + LTB + STB)",
                       pbt_cy + finance_cost_cy, total_equity_cy + ltb_cy + stb_cy,
                       pbt_py + finance_cost_py, total_equity_py + ltb_py + stb_py, true),

            calc_ratio("Debtor Turnover Ratio", "Revenue from Ops / Closing Debtors",
                       revenue_ops_cy, debtors_cy,
                       revenue_ops_py, debtors_py)
        };

        json output = json::array();
        for (const auto& r : ratios) {
            output.push_back({
                {"key", r.key},
                {"formula", r.formula},
                {"cy", r.cy},
                {"py", r.py},
                {"change", r.change},
                {"display_change", r.display_change},
                {"flag", r.flag}
            });
        }

        g_ratio_result = output.dump();
        return g_ratio_result.c_str();

    } catch (const std::exception& e) {
        json error = json::array();
        g_ratio_result = error.dump();
        return g_ratio_result.c_str();
    }
}

FAR_EXPORT const char* compute_ratios_from_raw(const char* bs_json, const char* pl_json) {
    try {
        json bs_rows = json::parse(bs_json);
        json pl_rows = json::parse(pl_json);
        
        json bs = json::object();
        json pl = json::object();
        
        double cy=0, py=0;
        
        find_value(bs_rows, "total equity", cy, py, true);
        bs["total_equity_cy"] = cy; bs["total_equity_py"] = py;
        
        find_value(bs_rows, "total current assets", cy, py);
        bs["total_ca_cy"] = cy; bs["total_ca_py"] = py;
        
        find_value(bs_rows, "total current liabilities", cy, py);
        bs["total_cl_cy"] = cy; bs["total_cl_py"] = py;
        
        find_value(bs_rows, "total non-current liabilities", cy, py);
        if (cy == 0 && py == 0) find_value(bs_rows, "total non current liabilities", cy, py);
        bs["total_ncl_cy"] = cy; bs["total_ncl_py"] = py;
        
        double ltb_pref_cy, ltb_pref_py, ltb_lease_cy, ltb_lease_py;
        find_liability_by_note(bs_rows, "11", false, ltb_pref_cy, ltb_pref_py);
        find_liability_by_note(bs_rows, "12", false, ltb_lease_cy, ltb_lease_py);
        bs["ltb_cy"] = ltb_pref_cy + ltb_lease_cy;
        bs["ltb_py"] = ltb_pref_py + ltb_lease_py;
        
        double stb_pref_cy, stb_pref_py, stb_lease_cy, stb_lease_py;
        find_liability_by_note(bs_rows, "11", true, stb_pref_cy, stb_pref_py);
        find_liability_by_note(bs_rows, "12", true, stb_lease_cy, stb_lease_py);
        bs["stb_cy"] = stb_pref_cy + stb_lease_cy;
        bs["stb_py"] = stb_pref_py + stb_lease_py;
        
        find_value(bs_rows, "trade receivables", cy, py);
        bs["debtors_cy"] = cy; bs["debtors_py"] = py;
        
        // PL
        find_value(pl_rows, "software services", cy, py);
        if (cy == 0 && py == 0) find_value(pl_rows, "revenue from operations", cy, py);
        pl["revenue_ops_cy"] = cy; pl["revenue_ops_py"] = py;
        
        find_value(pl_rows, "total income", cy, py);
        pl["total_revenue_cy"] = cy; pl["total_revenue_py"] = py;
        
        find_value(pl_rows, "profit", cy, py);
        for (const auto& row : pl_rows) {
            std::string p = to_lower(row.value("particulars", ""));
            if (p.find("profit") != std::string::npos && p.find("before") != std::string::npos && p.find("tax") != std::string::npos) {
                cy = row.value("cy", 0.0);
                py = row.value("py", 0.0);
                break;
            }
        }
        pl["pbt_cy"] = cy; pl["pbt_py"] = py;
        
        find_value(pl_rows, "finance charge", cy, py);
        if (cy == 0 && py == 0) find_value(pl_rows, "finance cost", cy, py);
        pl["finance_cost_cy"] = cy; pl["finance_cost_py"] = py;
        
        std::string bs_str = bs.dump();
        std::string pl_str = pl.dump();
        
        const char* ratios_cstr = compute_ratios(bs_str.c_str(), pl_str.c_str());
        json ratios_json = json::parse(ratios_cstr);
        
        json final_result = json::object();
        final_result["ratios"] = ratios_json;
        final_result["bs_summary"] = bs;
        final_result["pl_summary"] = pl;
        
        g_ratio_result = final_result.dump();
        return g_ratio_result.c_str();
        
    } catch (const std::exception& e) {
        json error = json::object();
        error["error"] = e.what();
        g_ratio_result = error.dump();
        return g_ratio_result.c_str();
    }
}

} // extern "C"
