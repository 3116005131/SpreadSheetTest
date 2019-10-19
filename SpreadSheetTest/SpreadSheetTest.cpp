// SpreadSheetTest.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include "pch.h"

#include <iostream>

#include"SpreadSheet.h"

#include<typeinfo>

int main()

{

	CSpreadSheet sheet("sheetfilename.xls", "sheetname",true);

	CStringArray headers, row,Rows;

	headers.Add("str1");

	headers.Add("str2");

	headers.Add("str3");

	row.Add("hello");

	row.Add("world");

	row.Add("!");



	sheet.BeginTransaction();

	//似乎必须要有headers，没有headers的话好像就没有输出文件。。。

	sheet.AddHeaders(headers);

	sheet.AddRow(row);

	sheet.Commit();

	//sheet.Convert(";");
	printf("表格有%ld行！\n", sheet.GetTotalRows());
	for(int i=1;i<=sheet.GetTotalRows();i++)
	{
		sheet.ReadRow(Rows, i);
		printf("第%d行，有%d个元素\n", i,Rows.GetSize());
		for (int j = 1; j <= Rows.GetSize(); j++)
		{
			if (j != Rows.GetSize())
			{
				printf("i=%d,j=%d\t", i, j);
				printf(Rows.GetAt(j-1));
				printf("\t");
			}
			else
			{
				printf("i=%d,j=%d\t", i, j);
				printf(Rows.GetAt(j - 1));
				printf("\n");
			}
		}
	}
}


// 运行程序: Ctrl + F5 或调试 >“开始执行(不调试)”菜单
// 调试程序: F5 或调试 >“开始调试”菜单

// 入门提示: 
//   1. 使用解决方案资源管理器窗口添加/管理文件
//   2. 使用团队资源管理器窗口连接到源代码管理
//   3. 使用输出窗口查看生成输出和其他消息
//   4. 使用错误列表窗口查看错误
//   5. 转到“项目”>“添加新项”以创建新的代码文件，或转到“项目”>“添加现有项”以将现有代码文件添加到项目
//   6. 将来，若要再次打开此项目，请转到“文件”>“打开”>“项目”并选择 .sln 文件
