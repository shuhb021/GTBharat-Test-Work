/**
 * FAR Engine — Shared Header
 * Common declarations for the C++ computation engine.
 * Compiled as far_engine.dll for Windows.
 */

#ifndef FAR_ENGINE_H
#define FAR_ENGINE_H

#ifdef _WIN32
    #define FAR_EXPORT __declspec(dllexport)
#else
    #define FAR_EXPORT __attribute__((visibility("default")))
#endif

extern "C" {

/**
 * Compute absolute and percentage variances for each row.
 * Input JSON: [{"particulars": "...", "note": "...", "cy": 123, "py": 456}, ...]
 * Output JSON: same array with added fields: variance_abs, variance_pct, display_pct, flag
 */
FAR_EXPORT const char* compute_variances(const char* json_in);

/**
 * Compute all 8 financial ratios.
 * Input: bs_json = {"total_equity_cy": ..., "total_equity_py": ..., ...}
 *        pl_json = {"revenue_ops_cy": ..., "pbt_cy": ..., ...}
 * Output JSON: [{"key": "Debt Equity Ratio", "formula": "...", "cy": ..., "py": ..., "change": ..., "flag": true/false}, ...]
 */
FAR_EXPORT const char* compute_ratios(const char* bs_json, const char* pl_json);

/**
 * Compute all 8 financial ratios directly from raw parsed rows.
 * Input: bs_json = raw BS rows array, pl_json = raw PL rows array
 * Output JSON: {"ratios": [...], "summary": {"total_equity_cy": ...}}
 */
FAR_EXPORT const char* compute_ratios_from_raw(const char* bs_json, const char* pl_json);

/**
 * Process large datasets with parallel computation.
 * Same as compute_variances but uses std::async for parallelism.
 */
FAR_EXPORT const char* process_bulk_data(const char* json_in, int row_count);

/**
 * Free memory allocated by the above functions.
 */
FAR_EXPORT void free_result(const char* ptr);

} // extern "C"

#endif // FAR_ENGINE_H
