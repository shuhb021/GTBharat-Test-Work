#include "far_engine.h"
#include "nlohmann/json.hpp"
#include <cmath>
#include <string>
#include <sstream>
#include <vector>
#include <future>
#include <algorithm>

using json = nlohmann::json;

static thread_local std::string g_bulk_result;

// Process a chunk of rows
static json process_chunk(const json& chunk) {
    json output = json::array();

    for (const auto& row : chunk) {
        double cy = row.value("cy", 0.0);
        double py = row.value("py", 0.0);
        double variance_abs = cy - py;
        double variance_pct = 0.0;
        std::string display_pct = "0.0%";
        bool flag = false;

        if (py == 0.0) {
            if (cy != 0.0) {
                display_pct = "N/A";
                flag = true;
            }
        } else {
            variance_pct = (variance_abs / std::fabs(py)) * 100.0;
            std::ostringstream oss;
            oss.precision(1);
            oss << std::fixed << variance_pct << "%";
            display_pct = oss.str();
            flag = std::fabs(variance_pct) >= 10.0;
        }

        json result_row = row;
        result_row["variance_abs"] = std::round(variance_abs * 100.0) / 100.0;
        result_row["variance_pct"] = (py == 0.0 && cy != 0.0) ? json(nullptr) :
                                     json(std::round(variance_pct * 100.0) / 100.0);
        result_row["display_pct"] = display_pct;
        result_row["flag"] = flag;
        output.push_back(result_row);
    }

    return output;
}

extern "C" {

FAR_EXPORT const char* process_bulk_data(const char* json_in, int row_count) {
    try {
        json input = json::parse(json_in);
        row_count = input.size();

        // For small datasets, process sequentially
        if (row_count < 100) {
            json result = process_chunk(input);
            g_bulk_result = result.dump();
            return g_bulk_result.c_str();
        }

        // Split into chunks for parallel processing
        const int num_threads = std::min(4, (row_count + 99) / 100);
        const int chunk_size = (row_count + num_threads - 1) / num_threads;

        std::vector<std::future<json>> futures;

        for (int i = 0; i < num_threads; ++i) {
            int start = i * chunk_size;
            int end = std::min(start + chunk_size, row_count);

            if (start >= row_count) break;

            json chunk = json::array();
            for (int j = start; j < end; ++j) {
                chunk.push_back(input[j]);
            }

            futures.push_back(std::async(std::launch::async, process_chunk, chunk));
        }

        json output = json::array();
        for (auto& f : futures) {
            json chunk_result = f.get();
            for (const auto& row : chunk_result) {
                output.push_back(row);
            }
        }

        g_bulk_result = output.dump();
        return g_bulk_result.c_str();

    } catch (const std::exception& e) {
        json error = json::array();
        g_bulk_result = error.dump();
        return g_bulk_result.c_str();
    }
}

} 