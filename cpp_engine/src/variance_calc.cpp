/**
 * FAR Engine — Variance Calculator
 * Computes absolute and percentage variances for financial line items.
 */

#include "far_engine.h"
#include "nlohmann/json.hpp"
#include <cmath>
#include <string>
#include <sstream>

using json = nlohmann::json;

// Thread-local storage for result strings
static thread_local std::string g_variance_result;

extern "C" {

FAR_EXPORT const char* compute_variances(const char* json_in) {
    try {
        json input = json::parse(json_in);
        json output = json::array();

        for (auto& row : input) {
            double cy = row.value("cy", 0.0);
            double py = row.value("py", 0.0);

            double variance_abs = cy - py;
            double variance_pct = 0.0;
            std::string display_pct = "0.0%";
            bool flag = false;

            if (py == 0.0) {
                if (cy == 0.0) {
                    variance_pct = 0.0;
                    display_pct = "0.0%";
                    flag = false;
                } else {
                    variance_pct = 0.0; // infinity represented as null
                    display_pct = "N/A";
                    flag = true;
                }
            } else {
                variance_pct = (variance_abs / std::fabs(py)) * 100.0;
                
                // Format display string
                std::ostringstream oss;
                oss.precision(1);
                oss << std::fixed << variance_pct << "%";
                display_pct = oss.str();
                
                flag = std::fabs(variance_pct) >= 10.0;
            }

            // Clone the input row and add computed fields
            json result_row = row;
            result_row["variance_abs"] = std::round(variance_abs * 100.0) / 100.0;
            
            if (py == 0.0 && cy != 0.0) {
                result_row["variance_pct"] = nullptr; // null for infinity
            } else {
                result_row["variance_pct"] = std::round(variance_pct * 100.0) / 100.0;
            }
            
            result_row["display_pct"] = display_pct;
            result_row["flag"] = flag;

            output.push_back(result_row);
        }

        g_variance_result = output.dump();
        return g_variance_result.c_str();

    } catch (const std::exception& e) {
        json error = json::array();
        g_variance_result = error.dump();
        return g_variance_result.c_str();
    }
}

FAR_EXPORT void free_result(const char* ptr) {
    // No-op: we use thread_local std::string, no manual free needed
    (void)ptr;
}

} // extern "C"

extern "C" {
FAR_EXPORT void compute_variances_fast(FinRow* rows, int count) {
    for (int i = 0; i < count; ++i) {
        double cy = rows[i].cy;
        double py = rows[i].py;
        
        double variance_abs = cy - py;
        double variance_pct = 0.0;
        bool flag = false;
        
        if (py == 0.0) {
            if (cy != 0.0) {
                variance_pct = -999999.0; 
                flag = true;
            }
        } else {
            variance_pct = (variance_abs / std::fabs(py)) * 100.0;
            flag = std::fabs(variance_pct) >= 10.0;
        }
        
        rows[i].variance_abs = std::round(variance_abs * 100.0) / 100.0;
        if (variance_pct != -999999.0) {
            rows[i].variance_pct = std::round(variance_pct * 100.0) / 100.0;
        } else {
            rows[i].variance_pct = -999999.0;
        }
        rows[i].flag = flag;
    }
}
}
