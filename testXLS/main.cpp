#include <stdio.h>
#include "xlslib.h"
#include <map>

using namespace std;
using namespace xlslib_core;

class MatchExcel
{
public:

	MatchExcel()
		:maker(wb.GetFormulaFactory())
		, offsetLine(6)
	{
// 		textFmt = wb.xformat();
// 		titleRange->fontbold(BOLDNESS_BOLD);
// 		titleRange->fillstyle(FILL_SOLID);
// 		titleRange->fillfgcolor(CLR_GRAY40);
// 		titleRange->halign(HALIGN_CENTER);
// 		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
// 		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
// 		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
// 		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);


	};
	~MatchExcel() {};
	
	struct MatchNode
	{
		MatchNode(std::string _nodeName)
			:nodeName(_nodeName)
			, nodeCount(1)
		{
		};

		MatchNode(std::string _nodeName, uint32_t _nodeCount)
			:nodeName(_nodeName)
			, nodeCount(_nodeCount)
		{
		};

		std::string nodeName; //题目名称
		uint32_t nodeCount; //小题数
	};
	typedef std::list<MatchNode> MatchNodes;

	struct MatchMap
	{
		MatchMap()
			:stuCount(1)
			, isPlusNode(false)
		{};

		uint32_t stuCount; //学生总数
		MatchNodes nodeList; //题目信息

		bool isPlusNode; //附加题
	};

	std::string toString(uint32_t valInt)
	{
		char tmp[8] = { 0 };
		snprintf(tmp, 8, "%u", valInt);

		return std::string(tmp);
	};

	std::string toColChar(uint32_t col)
	{
		char tmp[2] = { 0 };
		tmp[0] = col;

		return std::string(tmp);
	};

	expression_node_t * buildFuncSum(worksheet* ws, /*uint32_t target_row, uint32_t target_col,*/
		uint32_t first_row, uint32_t first_col, uint32_t last_row, uint32_t last_col)
	{
		cell_t* lefttop = ws->FindCellOrMakeBlank(first_row, first_col);
		cell_t* rightbottom = ws->FindCellOrMakeBlank(last_row, last_col);

		expression_node_t *area = maker.area(*lefttop, *rightbottom, CELL_RELATIVE_A1, CELLOP_AS_REFER);
		expression_node_t *areas[1];
		areas[0] = area;
		expression_node_t *f = maker.f(FUNC_SUM, 1, areas, CELL_DEFAULT);
// 		ws->formula(target_row, target_col, f, true);

		return f;
	};

	void buildTitle(worksheet* ws, uint32_t &curRow, uint32_t &curCol, const MatchNodes &nodeList)
	{
		//题目表
		for (MatchNodes::const_iterator iter = nodeList.begin(); iter != nodeList.end(); iter++)
		{
			if (iter->nodeCount > 1)
			{
				ws->merge(curRow, curCol, curRow, curCol + iter->nodeCount - 1);
				ws->label(curRow, curCol, iter->nodeName);

				for (uint32_t i = 0; i < iter->nodeCount; i++)
				{
					ws->label(curRow + 1, curCol + i, toString(i + 1));
				}
			}
			else
			{
				ws->merge(curRow, curCol, curRow + 1, curCol);
				ws->label(curRow, curCol, iter->nodeName);
			}
			curCol += iter->nodeCount;
		}
	}

	void buildLoss(worksheet* ws, uint32_t &curRow, uint32_t &curCol, const MatchNodes &nodeList, uint32_t stuCount)
	{
		//题目表
		for (MatchNodes::const_iterator iter = nodeList.begin(); iter != nodeList.end(); iter++)
		{
			if (iter->nodeCount > 1)
			{
				for (uint32_t i = 0; i < iter->nodeCount; i++)
				{
					expression_node_t * f = buildFuncSum(ws, /*curRow, curCol + i,*/ 2, curCol + i, 2 + stuCount - 1, curCol + i);
					ws->formula(curRow, curCol + i, f, true);
				}

				ws->merge(curRow + 1, curCol, curRow + 1, curCol + iter->nodeCount - 1);


				expression_node_t * f = buildFuncSum(ws, /*curRow + 1, curCol,*/ curRow, curCol, curRow, curCol + iter->nodeCount - 1);
				ws->formula(curRow + 1, curCol, f, true);

			}
			else
			{
				ws->merge(curRow, curCol, curRow + 1, curCol);
				expression_node_t * f = buildFuncSum(ws, /*curRow, curCol,*/ 2, curCol, 2 + stuCount - 1, curCol);
				ws->formula(curRow, curCol, f, true);
			}
			curCol += iter->nodeCount;
		}
	}

	void actTitle(worksheet* ws, unsigned32_t row1, unsigned32_t col1,
		unsigned32_t row2, unsigned32_t col2)
	{
		range* titleRange = ws->rangegroup(row1, col1, row2, col2);

		titleRange->fontbold(BOLDNESS_BOLD);
		titleRange->fillstyle(FILL_SOLID);
		titleRange->fillfgcolor((color_name_t)40);
		titleRange->halign(HALIGN_CENTER);
		titleRange->valign(VALIGN_CENTER);
		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);

		titleRange->locked(true);
	}

	void actFunc(worksheet* ws, unsigned32_t row1, unsigned32_t col1,
		unsigned32_t row2, unsigned32_t col2)
	{
		range* titleRange = ws->rangegroup(row1, col1, row2, col2);
		titleRange->fillstyle(FILL_SOLID);
		titleRange->fillfgcolor((color_name_t)17);
		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);

		titleRange->locked(true);
	}

	void actStu(worksheet* ws, unsigned32_t row1, unsigned32_t col1,
		unsigned32_t row2, unsigned32_t col2, unsigned32_t sumColNo)
	{
		bool isBule = true;

		for (unsigned32_t row = row1; row <= row2; row++)
		{
			range* titleRange = ws->rangegroup(row, col1, row, col2);
			if (isBule)
			{
				titleRange->fillstyle(FILL_SOLID);
				titleRange->fillfgcolor((color_name_t)28);
			}
			titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
			titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
			titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
			titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);
			titleRange->locked(false);

			isBule = !isBule;
		}

		range* stuNo = ws->rangegroup(row1, 0, row2, 0);
		stuNo->fillstyle(FILL_SOLID);
		stuNo->fillfgcolor((color_name_t)28);
		
		range* sumCol = ws->rangegroup(row1, sumColNo, row2, sumColNo);
		sumCol->locked(true);
		if (sumColNo == col2)
		{
			sumCol->fillstyle(FILL_SOLID);
			sumCol->fillfgcolor((color_name_t)28);
		}
		else
		{
			range* lastCol = ws->rangegroup(row1, col2, row2, col2);
			lastCol->locked(true);
			lastCol->fillstyle(FILL_SOLID);
			lastCol->fillfgcolor((color_name_t)28);
		}
	}

	void actEdit(worksheet* ws, unsigned32_t row1, unsigned32_t col1,
		unsigned32_t row2, unsigned32_t col2)
	{
		bool isBule = true;

		for (unsigned32_t row = row1; row <= row2; row++)
		{
			range* titleRange = ws->rangegroup(row, col1, row, col2);
			if (isBule)
			{
				titleRange->fillstyle(FILL_SOLID);
				titleRange->fillfgcolor((color_name_t)28);
			}
			titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
			titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
			titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
			titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);
			titleRange->locked(false);

			isBule = !isBule;
		}
	}

	void inputExcel(const MatchMap & inData)
	{
		uint32_t SumCol; //总分列

		uint32_t lossRow; //失分行
		

		worksheet* ws = wb.sheet(L"统分表");
		
		ws->defaultColwidth(8);
		ws->defaultRowHeight(18);

		wb.setColor(196, 215, 155, 9); //title
		wb.setColor(250, 191, 143, 10); //func
		wb.setColor(184, 204, 228, 11); //stu


		uint32_t curCol = 0;
		uint32_t curRow = 0;

		//学号
		ws->merge(curRow, curCol, curRow + 1, curCol);
		ws->label(curRow, curCol, L"学号");
		curCol++;

		//姓名
		ws->merge(curRow, curCol, curRow + 1, curCol);
		ws->label(curRow, curCol, L"姓名");
		curCol++;

		buildTitle(ws, curRow, curCol, inData.nodeList);

		//总分
		SumCol = curCol;
		ws->merge(curRow, curCol, curRow + 1, curCol);
		ws->label(curRow, curCol, L"总分");

		//附加题
		if (inData.isPlusNode)
		{
			curCol++;
			ws->merge(curRow, curCol, curRow + 1, curCol);
			ws->label(curRow, curCol, L"附加题");
		}

		actTitle(ws, curRow, 0, curRow + 1, curCol);

		curRow += 2;

		//姓名表
		
		cell_t* totalScore = ws->FindCellOrMakeBlank(4 + offsetLine + inData.stuCount, SumCol);
		for (uint32_t i = curRow; i < curRow + inData.stuCount; i ++)
		{
			expression_node_t * sumLoss = buildFuncSum(ws, i, 2, i, SumCol - 1);
			expression_node_t * score = maker.op(OP_SUB, maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), sumLoss);
			ws->formula(i, SumCol, score, true);
		}

		actStu(ws, curRow, 0, curRow + inData.stuCount - 1, curCol, SumCol);

		curRow += inData.stuCount;

		//失分
		lossRow = curRow;
		curCol = 0;
		ws->merge(curRow, curCol, curRow + 1, curCol + 1);
		ws->label(curRow, curCol, L"失分");
		actTitle(ws, curRow, curCol, curRow + 1, curCol + 1);

		curCol = 2;


		//失分统计
		buildLoss(ws, curRow, curCol, inData.nodeList, inData.stuCount);

		//失分总分
		ws->merge(curRow, curCol, curRow + 1, curCol);
		expression_node_t * losFunc = buildFuncSum(ws, /*curRow, curCol,*/ curRow, 2, curRow, curCol - 1);
		ws->formula(curRow, curCol, losFunc, true);

		actFunc(ws, curRow, 2, curRow + 1, curCol);

		curCol += 1;

		ws->rowheight(0, 20 * 20);
		ws->rowheight(1, 20 * 20);
		for (uint32_t i = 2; i <= curRow; i++)
		{
			ws->rowheight(i, 18 * 20);
		}
		ws->colwidth(0, 6*256);
		ws->colwidth(1, 12*256);


		//小题单项分数
		{
			curRow += offsetLine;
			curCol = 1;
			ws->label(curRow, 1, L"大题");
			ws->label(curRow + 1, 1, L"小题");
			ws->label(curRow + 2, 1, L"单项总分");

			curCol += 1;
			buildTitle(ws, curRow, curCol, inData.nodeList);

			actEdit(ws, curRow + 2, 2, curRow + 2, curCol);

			//试卷总分
			SumCol = curCol;
			ws->merge(curRow, curCol, curRow + 1, curCol);
			ws->label(curRow, curCol, L"总分");

			//附加题总分
			if (inData.isPlusNode)
			{
				curCol++;
				ws->merge(curRow, curCol, curRow + 1, curCol);
				ws->label(curRow, curCol, L"附加题");

				actEdit(ws, curRow + 2, 2, curRow + 2, curCol);
			}

			actTitle(ws, curRow, 1, curRow + 2, 1);
			actTitle(ws, curRow, 1, curRow + 1, curCol);


			//总分
			curRow += 2;
			curCol = SumCol;
			expression_node_t * totalFunc = buildFuncSum(ws, curRow, 2, curRow, curCol - 1);
			ws->formula(curRow, curCol, totalFunc, true);

			actFunc(ws, curRow, curCol, curRow, curCol);
		}

		//丢分统计
		{
			curRow += 2;
			curCol = 1;
			ws->label(curRow, 1, L"应得分");
			ws->label(curRow + 1, 1, L"实得分");
			ws->label(curRow + 2, 1, L"得分率");

			actTitle(ws, curRow, 1, curRow + 2, 1);

			curCol += 1;
			//单项
			for (MatchNodes::const_iterator iter = inData.nodeList.begin(); iter != inData.nodeList.end(); iter++)
			{
				for (uint32_t i = 0; i < iter->nodeCount; i++)
				{
					//应得分
					cell_t *oneTotal = ws->FindCellOrMakeBlank(curRow - 2, curCol + i);
					expression_node_t *totalFunc = maker.op(xlslib_core::OP_MUL, maker.integer((signed32_t)inData.stuCount), maker.cell(*oneTotal, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

					ws->formula(curRow, curCol+i, totalFunc, true);

					//实得分
					cell_t *totalScore = ws->FindCellOrMakeBlank(curRow, curCol + i);
					cell_t *totalLoss = ws->FindCellOrMakeBlank(lossRow, curCol + i);
					expression_node_t *actScore = maker.op(xlslib_core::OP_SUB, maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*totalLoss, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

					ws->formula(curRow + 1, curCol + i, actScore, true);

					//得分率
					cell_t *realScore = ws->FindCellOrMakeBlank(curRow + 1, curCol + i);
					expression_node_t *scorePercent = maker.op(xlslib_core::OP_DIV, maker.cell(*realScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
					
					xf_t* sxf1 = wb.xformat();
					sxf1->SetFormat(FMT_PERCENT2);
					ws->formula(curRow + 2, curCol + i, scorePercent, true, sxf1);
				}
				curCol += iter->nodeCount;
			}

			//总分
			{
				//应得分
				cell_t *oneTotal = ws->FindCellOrMakeBlank(curRow - 2, curCol);
				expression_node_t *totalFunc = maker.op(xlslib_core::OP_MUL, maker.integer((signed32_t)inData.stuCount), maker.cell(*oneTotal, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

				ws->formula(curRow, curCol, totalFunc, true);

				//实得分
				cell_t *totalScore = ws->FindCellOrMakeBlank(curRow, curCol);
				cell_t *totalLoss = ws->FindCellOrMakeBlank(lossRow, curCol);
				expression_node_t *actScore = maker.op(xlslib_core::OP_SUB, maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*totalLoss, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

				ws->formula(curRow + 1, curCol, actScore, true);

				//得分率
				cell_t *realScore = ws->FindCellOrMakeBlank(curRow + 1, curCol);
				expression_node_t *scorePercent = maker.op(xlslib_core::OP_DIV, maker.cell(*realScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

				xf_t* sxf1 = wb.xformat();
				sxf1->SetFormat(FMT_PERCENT2);
				ws->formula(curRow + 2, curCol, scorePercent, true, sxf1);

			}

			//附加分
			if (inData.isPlusNode)
			{
				curCol += 1;

				//应得分
				cell_t *oneTotal = ws->FindCellOrMakeBlank(curRow - 2, curCol);
				expression_node_t *totalFunc = maker.op(xlslib_core::OP_MUL, maker.integer((signed32_t)inData.stuCount), maker.cell(*oneTotal, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

				ws->formula(curRow, curCol, totalFunc, true);

				//实得分
				cell_t *totalScore = ws->FindCellOrMakeBlank(curRow, curCol);
				expression_node_t *actScore = buildFuncSum(ws, 2, curCol, 2 + inData.stuCount - 1, curCol);

				ws->formula(curRow + 1, curCol, actScore, true);

				//得分率
				cell_t *realScore = ws->FindCellOrMakeBlank(curRow + 1, curCol);
				expression_node_t *scorePercent = maker.op(xlslib_core::OP_DIV, maker.cell(*realScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

				xf_t* sxf1 = wb.xformat();
				sxf1->SetFormat(FMT_PERCENT2);
				ws->formula(curRow + 2, curCol, scorePercent, true, sxf1);
			}

			actFunc(ws, curRow, 2, curRow + 2, curCol);

			for (uint32_t i = SumCol - 4; i < SumCol + 3; i++)
			{
				ws->rowheight(i, 18 * 20);
			}
		}
	}

	void outputExcel(const std::string &outFilePath)
	{
		wb.Dump(outFilePath);
	};


private:
	workbook wb;
	expression_node_factory_t& maker;

	uint32_t offsetLine;
};



void test()
{
	MatchExcel newExcel;

	MatchExcel::MatchMap inData;
	inData.stuCount = 61;
	inData.nodeList.push_back(MatchExcel::MatchExcel::MatchNode("First", 1));
	inData.nodeList.push_back(MatchExcel::MatchNode("Second", 2));
	inData.nodeList.push_back(MatchExcel::MatchNode("Third", 3));
	inData.nodeList.push_back(MatchExcel::MatchNode("Forth", 5));
	inData.isPlusNode = false;

	newExcel.inputExcel(inData);
	newExcel.outputExcel("test.xls");
}

int main()
{
	test();

	return 0;
	int a = 0;

	/////
	workbook wb;

	worksheet* sh = wb.sheet("NUMBERS");
	expression_node_factory_t& maker = wb.GetFormulaFactory();

	const unsigned int len = 4;
	unsigned int row = 1;
	unsigned int formula_col = len + 1;

	sh->label(0, formula_col, "FORMULAS");

	// SUM(cell, cell, cell, cell)
	expression_node_t *cells[len];
	for (unsigned int i = 0; i < len; ++i) {
		cell_t *c = sh->number(row, i, 1 + i);
		cells[i] = maker.cell(*c, CELL_RELATIVE_A1, CELLOP_AS_VALUE);
	}
	{
		expression_node_t *f = maker.f(FUNC_SUM, len, cells, CELL_DEFAULT); // CELL_DEFAULT CELLOP_AS_ARRAY
		sh->formula(row, formula_col, f, true);
	}

	// SUM(cell:cell)
	++row;
	cell_t *real_cells[len];
	for (unsigned int i = 0; i < len; ++i) {
		real_cells[i] = sh->number(row, i, (1 + i)*row);
		//cells[i] = maker.cell(*c, CELL_RELATIVE_A1, CELLOP_AS_VALUE);
	}
	{
		expression_node_t *area = maker.area((cell_t&)*(real_cells[0]), (cell_t&)*(real_cells[len - 1]), CELL_RELATIVE_A1, CELLOP_AS_REFER);
		expression_node_t *areas[1];
		areas[0] = area;
		expression_node_t *f = maker.f(FUNC_SUM, 1, areas, CELL_DEFAULT);
		sh->formula(row, formula_col, f, true);
	}
	wb.Dump("workbook.xls");

	return 0;
}