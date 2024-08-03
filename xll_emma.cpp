// xll_emma.cpp - Get EMMA data.
// https://emma.msrb.org/TreasuryData/GetTreasuryDailyYieldCurve?curveDate=mm/dd/yyyy
// https://emma.msrb.org/ICEData/GetICEDailyYieldCurve?curveDate=
// https://emma.msrb.org/TradeData/MostRecentTrades
// https://emma.msrb.org/TradeData/GetMostRecentTrades?
// https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/2023/all?type=daily_treasury_yield_curve&field_tdr_date_value=2023&page&_format=csv

#include "fms_sqlite/fms_sqlite.h"
#include "xll24/xll.h"

#define CATEGORY L"EMMA"

using namespace xll;

// TODO: more curves
const OPER TreasuryDailyYieldCurve("https://emma.msrb.org/TreasuryData/GetTreasuryDailyYieldCurve?curveDate=");
const OPER ICEDailyYieldCurve("https://emma.msrb.org/ICEData/GetICEDailyYieldCurve?curveDate=");

std::map<std::string, OPER> curve_map = {
	{"Treasury", TreasuryDailyYieldCurve},
	{"ICE", ICEDailyYieldCurve},
};

XLL_CONST(LPOPER, EMMA_ENUM_TREASURY, (LPOPER)&TreasuryDailyYieldCurve, "US Treasury Constant Maturity Treasury rates.", "EMMA", "");
XLL_CONST(LPOPER, EMMA_ENUM_ICE, (LPOPER)&ICEDailyYieldCurve, "EMMA ICE daily yield curve.", "EMMA", "");

// in memory database
sqlite::db db("");
sqlite::db emma_db();

int create_db()
{
	sqlite::stmt stmt(::db);

	ensure(SQLITE_DONE == stmt.exec("DROP TABLE IF EXISTS emma"));

	return stmt.exec("CREATE TABLE emma (curve TEXT, date FLOAT, data JSON)"));
}
int open_emma_db()
{
	emma_db.open("emma.db");
	sqlite::stmt stmt(emma_db);

	return emma_db.exec(
		"CREATE TABLE IF NOT EXISTS data("
		"curve TEXT, date FLOAT, year FLOAT, rate FLOAT),"
		"PRIMARY KEY(curve, date, year)"
	);
}

Auto<Open> xao_emma_db([] {
	try {
		create_db();

		// local database
		OPER module = Excel(xlGetName);
		auto view = xll::view(module);
		auto i = view.find_last_of(L"\\/");
		auto dir = view.substr(0, i + 1);
		ensure(SetCurrentDirectory(std::string(dir).c_str()));

		ensure(SQLITE_OK == open_emma_db());
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return 1;
});

int insert_curve_row(const std::string_view curve, double date, std::wstring_view data)
{
	try {
		ensure(curve_map.contains(std::string(curve)));

		sqlite::stmt stmt(::db);
		stmt.prepare("INSERT INTO emma VALUES(?, ?, ?)");
		stmt.bind(1, curve);
		stmt.bind(2, date);
		stmt.bind(3, data);
		ensure(SQLITE_DONE == stmt.step());
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

int insert_curve_date(const std::string_view name, double date)
{
	try {
		const auto i = curve_map.find(std::string(name));
		ensure(i != curve_map.end());
		const OPER& url = i->second;
		OPER data = Excel(xlfWebservice, url & Excel(xlfText, date, L"mm/dd/yyyy"));
		// {"Series":[{"Points":[{"X":"1","Y":"2.903"},...]]
		if (!data || view(data).starts_with(L"{\"Series\":[]")) {
			return 0; // no data
		}

		insert_curve_row(name, date, view(data));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

// {"Series":[{"Points":[{"X":"1","Y":"2.896"},...,{"X":"30","Y":"3.660"}]}]}
constexpr const char* sql_select 
	= "SELECT json_extract(value, '$.X'), json_extract(value, '$.Y') "
	  "FROM emma, json_each(json_extract(data, '$.Series[0].Points')) "
	  "WHERE curve = ? AND date = ?";

inline FPX get_curve_points(const char* name, double date)
{
	FPX result;

	try {
		sqlite::stmt stmt(::db);
		stmt.prepare(sql_select);
		stmt.bind(1, name);
		stmt.bind(2, date);

		while (SQLITE_ROW == stmt.step()) {
			result.vstack(FPX({ stmt[0].as_float(), stmt[1].as_float()/100 }));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return result;
}

inline FPX get_insert_curve_points(const char* name, double date)
{
	FPX result;

	try {
		result = get_curve_points(name, date);
		if (!result.size()) {
			if (insert_curve_date(name, date)) {
				result = get_curve_points(name, date);
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return result;
}

AddIn xai_emma_curve(
	Function(XLL_FP, "xll_emma_curve", "EMMA")
	.Arguments({
		Arg(XLL_CSTRING4, "name", "is either \"Treasury\" or \"ICE\"."),
		Arg(XLL_DOUBLE, "date", "is the date of the curve. Default is previous business day."),
		})
	.Category(CATEGORY)
	.FunctionHelp("EMMA curves as two row array of years and par coupon rates.")
);
FP12* WINAPI xll_emma_curve(const char* name, double date)
{
#pragma XLLEXPORT
	static FPX result;

	try {
		if (!date) {
			date = asNum(Excel(xlfWorkday, Excel(xlfToday), -1));
		}

		result = get_insert_curve_points(name, date);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return nullptr;
	}

	return result.get();
}