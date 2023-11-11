#include <iostream>
#include <Windows.h>
#include "MiniExcel.cpp"
using namespace std;
int main()
{
    system("cls");
    MiniExcel obj;
    obj.printCell(0,0,10);
    while (true)
    {
        if (GetAsyncKeyState('S'))
        {
            obj.moveDown();
        }
        if (GetAsyncKeyState('W'))
        {
            obj.moveUp();
        }
        if (GetAsyncKeyState('D'))
        {
            obj.moveRight();
        }
        if (GetAsyncKeyState('A'))
        {
            obj.moveLeft();
        }
        if (GetAsyncKeyState('I'))
        {
            obj.goToRowCol(12, 82);
            string value;
            cout <<"Enter value: ";
            cin >> value;
            obj.insertValue(value);
        }
        if (GetAsyncKeyState(VK_UP))
        {
            obj.insertRowAbove();
            obj.printGrid();
            obj.printData();
        }
        if (GetAsyncKeyState(VK_DOWN))
        {
            obj.insertRowBelow();
            obj.printGrid();
            obj.printData();
        }
        if (GetAsyncKeyState(VK_RIGHT))
        {
            obj.insertColumnToRight();
            obj.printGrid();
            obj.printData();
        }
        if (GetAsyncKeyState(VK_LEFT))
        {
            obj.insertColumnToLeft();
            obj.printGrid();
            obj.printData();
        }
        if (GetAsyncKeyState('L'))
        {
            obj.insertCellByRightShift();
            obj.printGrid();
            obj.printData();
        }
        if (GetAsyncKeyState('K'))
        {
            obj.insertCellByLeftShift();
            obj.printGrid();
            obj.printData();
        }
        if (GetAsyncKeyState('M'))
        {
            obj.insertCellByDownShift();
            obj.printGrid();
            obj.printData();
        }
        if (GetAsyncKeyState('O'))
        {
            obj.insertCellByUpShift();
            obj.printGrid();
            obj.printData();
        }
        Sleep(100);
    }
    // obj.moveCursor();
    // obj.printGrid();
    // obj.printCell(0,0,10);
    // obj.printCell(0,0,3);
    // obj.printCell(0,1,3);
    //  obj.printCell(1,0,3);
    //  obj.printCell(1,1,3);

}