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

using json = nlohmann::json;

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
                       pbt_cy, total_equity_cy,
                       pbt_py, total_equity_py, true),

            calc_ratio("Return on Capital Employed (ROCE)", "(PBT - Finance Cost) / (TE + LTB + STB)",
                       pbt_cy - finance_cost_cy, total_equity_cy + ltb_cy + stb_cy,
                       pbt_py - finance_cost_py, total_equity_py + ltb_py + stb_py, true),

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

} // extern "C"
