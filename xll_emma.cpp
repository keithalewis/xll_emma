﻿// xll_emma.cpp - Get EMMA data.
// https://emma.msrb.org/TreasuryData/GetTreasuryDailyYieldCurve?curveDate=mm/dd/yyyy
// https://emma.msrb.org/ICEData/GetICEDailyYieldCurve?curveDate=
// https://emma.msrb.org/TradeData/MostRecentTrades
// https://emma.msrb.org/TradeData/GetMostRecentTrades?
// https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/2023/all?type=daily_treasury_yield_curve&field_tdr_date_value=2023&page&_format=csv

#define CATEGORY L"EMMA"

//#include <string_view>
#include "fms_sqlite/fms_sqlite.h"
#include "xll24/include/xll.h"

using namespace xll;

#define EMMA_HASH(x) L#x
#define EMMA_STRZ(x) EMMA_HASH(x)

// Curve info: Series.Name, Id, EarliestAvailableFilingDate, LatestAvailableFilingDate, Description
// Assumes Id is unique.
// https://emma.msrb.org/ToolsAndResources/BondWaveYieldCurve?daily=True
// Curves are stored as Source_Id in the database.
// Source, Id, Name, Desc
#define EMMA_CURVES(X) \
	X(Bloomberg, CAAA, "BVAL® AAA Callable Municipal Curve.", \
		"The BVAL® AAA Municipal Curves use dynamic real-time trades and contributed sources to reflect movement in the municipal market.") \
	X(Bloomberg, BVMB, "BVAL® AAA Municipal Curve.", \
		"The BVAL® AAA Municipal Curves use dynamic real-time trades and contributed sources to reflect movement in the municipal market.") \
	X(BondWave, BondWave, "BondWave AA QCurve.", \
		"The BondWave AA QCurve is a quantitatively derived yield curve built from executed trades offering full data transparency.") \
	X(ICE, ICE, "ICE US Municipal AAA Curve.", \
		"The ICE US Municipal AAA Yield Curve is produced continuously and used daily to apply intraday and end-of-day market moves to the majority of the investment grade municipal bond universe") \
	X(IHSMarkit, AAA, "IHS Markit Municipal Bond AAA Curve.", \
		"The IHS Markit Municipal Bond AAA Curve is a tax-exempt yield curve that consists of 5% General Obligation AAA debt, callable after 10 years.") \
	X(MBIS, MBIS, "MBIS AAA Municipal Curve.", \
		"The MBIS Municipal Benchmark Curve is a tax-exempt investment grade yield curve that is valued directly against pre- and post-trade market data provided by the MSRB.") \
	X(Tradeweb, Tradeweb, "Tradeweb AAA Municipal Yield Curve.", \
		"Tradeweb’s Ai-Price for Municipal Bonds addresses the challenge of price discovery by leveraging proprietary machine learning and data science combined with MSRB and Tradeweb proprietary data to price approximately one million municipal bonds at or near traded prices.") \
	X(Treasury, Treasury, "Treasury Yield Curve Rates.", \
		"U.S. Treasury Yield Curve Rates are commonly referred to as \"Constant Maturity Treasury\" rates, or CMTs.") \

#define EMMA_CURVE_URL(source) \
	L"https://emma.msrb.org/" source L"Data/Get" source L"DailyYieldCurve?curveDate="
#define EMMA_CURVE_TOPIC(source) \
	L"https://emma.msrb.org/ToolsAndResources/" source L"YieldCurve?daily=True"

// Known Id.
bool contains(std::wstring_view Id)
{
#define EMMA_CURVE_ENUM(source, id, name, desc) \
	if (Id == L#id) return true; 
	EMMA_CURVES(EMMA_CURVE_ENUM)
#undef EMMA_CURVE_ENUM

	return false;
}
// Source given Id
const wchar_t* source_id(std::wstring_view Id)
{
#define EMMA_CURVE_ENUM(source, id, name, desc) \
	if (Id == L#id) return L#source; 
	EMMA_CURVES(EMMA_CURVE_ENUM)
#undef EMMA_CURVE_ENUM

	return nullptr;
}
// Source url from Id
const wchar_t* url_id(std::wstring_view Id)
{
#define EMMA_CURVE_ENUM(source, id, name, desc) \
	if (Id == L#id) return EMMA_CURVE_URL(EMMA_HASH(source));
	EMMA_CURVES(EMMA_CURVE_ENUM)
#undef EMMA_CURVE_ENUM

	return nullptr;
}
// Help topic url from Id
const wchar_t* topic_id(std::wstring_view Id)
{
#define EMMA_CURVE_ENUM(source, id, name, desc) \
	if (Id == L#id) return EMMA_CURVE_TOPIC(EMMA_HASH(source));
	EMMA_CURVES(EMMA_CURVE_ENUM)
#undef EMMA_CURVE_ENUM

		return nullptr;
}

// Enums with storage.
#define EMMA_CURVE_ENUM(source, id, name, desc) \
	const OPER source##_##id##_ENUM = OPER(L#id);
EMMA_CURVES(EMMA_CURVE_ENUM)
#undef EMMA_CURVE_ENUM

#define EMMA_CURVE_ENUM(source, id, name, desc) \
	XLL_CONST(LPOPER, EMMA_##id, (LPOPER)&source##_##id##_ENUM, name, "EMMA Enum", EMMA_CURVE_TOPIC(EMMA_HASH(source)));
EMMA_CURVES(EMMA_CURVE_ENUM)
#undef EMMA_CURVE_ENUM

#define EMMA_CURVE_ENUM(source, id, name, desc) \
	OPER(EMMA_STRZ(id)),
OPER EMMA_Enum = OPER({
	EMMA_CURVES(EMMA_CURVE_ENUM)
});
#undef EMMA_CURVE_ENUM
XLL_CONST(LPOPER, EMMA_ENUM, (LPOPER)&EMMA_Enum, "EMMA curve enumeration.", "EMMA Enum", L"https://emma.msrb.org/ToolsAndResources/MarketIndicators");

// in add-in directory
sqlite::db emma_db;
// TODO: put in db subdirectory to prevent XLSTART from opening???

// Directory of add-in
const std::wstring& xll_dir()
{
	static std::wstring dir;

	if (dir.empty()) {
		OPER xll = Excel(xlGetName);
		auto v = xll::view(xll);
		dir = v.substr(0, 1 + v.find_last_of(L"\\/"));
		if (dir.ends_with(L"Debug\\")) {
			dir = dir.substr(0, dir.size() - 6);
		}
		else if (dir.ends_with(L"Release\\")) {
			dir = dir.substr(0, dir.size() - 8);
		}
		if (dir.ends_with(L"x64\\")) {
			dir = dir.substr(0, dir.size() - 4);
		}
	}

	return dir;
}

int create_emma_db()
{
	auto path = xll_dir() + L"emma.db";

	emma_db.open(path.c_str());
	sqlite::stmt stmt(emma_db);

	return stmt.exec(R"(
		CREATE TABLE IF NOT EXISTS curve(
		source_id TEXT, date FLOAT, year FLOAT, rate FLOAT, 
		PRIMARY KEY(source_id, date, year))
	)");

}
Auto<Open> xao_emma_db([] {
	try {
		ensure(SQLITE_DONE == create_emma_db());
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return 1;
});

//double EarliestAvailableFilingDate(const OPER& data) {}
//double LatestAvailableFilingDate(const OPER& data) {}

// Most recent date EMMA has for id prior or equal to date.
double emma_id_date(const std::wstring_view id, double date)
{
	double id_date = -1;

	sqlite::stmt stmt(emma_db);
	stmt.prepare("SELECT date "
		"FROM curve "
		"WHERE source_id = ? AND date <= ? "
		"ORDER BY date DESC "
		"LIMIT 1");
	stmt.bind(1, id);
	stmt.bind(2, date);
	if (SQLITE_ROW == stmt.step()) {
		id_date = stmt[0].as_float();
	}

	return id_date;
}

// Insert curve given id and date. Return date of most recent data.
double insert_id_date(const std::wstring_view id, double date)
{
	// TODO: async
	OPER url = OPER(url_id(id));
	OPER data = Excel(xlfWebservice, url & Excel(xlfText, date, L"mm/dd/yyyy"));
	// {"Series":[{"Points":[{"X":"1","Y":"2.903"},...]]
	while (!data || view(data).starts_with(L"{\"Series\":[]")) {
		date = asNum(Excel(xlfWorkday, date, -1));
		if (date < asNum(Excel(xlfDate, 2015, 1, 1))) { // avoid infinite loop
			return -1;
		}
		if (date == emma_id_date(id, date)) {
			return date;
		}
		data = Excel(xlfWebservice, url & Excel(xlfText, date, L"mm/dd/yyyy"));
	}

	const char* sql = R"(
INSERT INTO curve(source_id, date, year, rate)
WITH p AS (SELECT :data AS data)
SELECT
   iif(instr(json_extract(series.value, '$.Id'), " "),
		substr(json_extract(series.value, '$.Id'), 0, instr(json_extract(series.value, '$.Id'), " ")),
		json_extract(series.value, '$.Id')),
	:date as date,
    json_extract(point.value, '$.X') AS year,
    json_extract(point.value, '$.Y') AS rate
FROM p, json_each(p.data, '$.Series') AS series,
    json_each(series.value, '$.Points') AS point
)";
	sqlite::stmt stmt(emma_db);
	stmt.prepare(sql);
	stmt.bind(":data", view(data));
	stmt.bind(":date", date);

	ensure(SQLITE_DONE == stmt.step());

	return date;
}

inline FPX get_curve_points(std::wstring_view id, double date)
{
	FPX result;

	sqlite::stmt stmt(emma_db);
	stmt.prepare(
		"SELECT year, rate FROM curve "
		"WHERE source_id = ? and date = ? "
		"ORDER BY year"
	);
	stmt.bind(1, id);
	stmt.bind(2, date);

	while (SQLITE_ROW == stmt.step()) {
		result.append(stmt[0].as_float())
			  .append(stmt[1].as_float() / 100);
	}
	if (result.size()) {
		result.resize(result.size()/2, 2);
	}

	return result;
}

inline FPX get_insert_curve_points(std::wstring_view curve, double date)
{
	FPX result;

	result = get_curve_points(curve, date);

	if (!result.size()) {
		date = insert_id_date(curve, date);
		ensure(date > 0);
		result = get_curve_points(curve, date);
	}

	return result;
}

AddIn xai_emma_source(
	Function(XLL_CSTRING, "xll_emma_source", "EMMA.SOURCE")
	.Arguments({
		Arg(XLL_CSTRING, "id", "is a value from the EMMA_ENUM() enumeration."),
		})
		.Category(CATEGORY)
	.FunctionHelp("EMMA data source for Id.")
);
const wchar_t* WINAPI xll_emma_source(const wchar_t* id)
{
#pragma XLLEXPORT
	try {
		ensure(contains(id));

		ensure(id = source_id(id));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return id;
}

AddIn xai_emma_url(
	Function(XLL_CSTRING, "xll_emma_url", "EMMA.URL")
	.Arguments({
		Arg(XLL_CSTRING, "id", "is a value from the EMMA_ENUM() enumeration."),
		})
		.Category(CATEGORY)
	.FunctionHelp("EMMA source URL from Id.")
);
const wchar_t* WINAPI xll_emma_url(const wchar_t* id)
{
#pragma XLLEXPORT
	try {
		ensure(contains(id));

		ensure(id = url_id(id));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return id;
}

AddIn xai_emma_help(
	Function(XLL_CSTRING, "xll_emma_help", "EMMA.HELP")
	.Arguments({
		Arg(XLL_CSTRING, "id", "is a value from the EMMA_ENUM() enumeration."),
		})
		.Category(CATEGORY)
	.FunctionHelp("EMMA source help from Id.")
);
const wchar_t* WINAPI xll_emma_help(const wchar_t* id)
{
#pragma XLLEXPORT
	try {
		ensure(contains(id));

		ensure(id = topic_id(id));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return id;
}

AddIn xai_emma_curve(
	Function(XLL_FP, "xll_emma_curve", "EMMA")
	.Arguments({
		Arg(XLL_CSTRING, "id", "is a value from the EMMA_ENUM() enumeration."),
		Arg(XLL_DOUBLE, "date", "is the date of the curve. Default is the most recent data available."),
		})
	.Category(CATEGORY)
	.FunctionHelp("EMMA curves as two row array of years and par coupon rates.")
);
FP12* WINAPI xll_emma_curve(const wchar_t* id, double date)
{
#pragma XLLEXPORT
	static FPX result;

	try {
		ensure(contains(id));

		if (!date) {
			date = asNum(Excel(xlfWorkday, Excel(xlfToday), -1));
		}

		result = get_insert_curve_points(id, date);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return nullptr;
	}

	return result.get();
}

AddIn xai_emma_date(
	Function(XLL_DOUBLE, "xll_emma_date", "EMMA.DATE")
	.Arguments({
		Arg(XLL_CSTRING, "id", "is a value from the EMMA_ENUM() enumeration."),
		Arg(XLL_DOUBLE, "date", "is the date of the date. Default is the most recent data available."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return most recent date EMMA has for id.")
);
double WINAPI xll_emma_date(const wchar_t* id, double date)
{
#pragma XLLEXPORT
	try {
		ensure(contains(id));

		if (!date) {
			date = asNum(Excel(xlfToday));
		}
		date = emma_id_date(id, date);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return date;
}