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
        char c =_getch();
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
        if (GetAsyncKeyState('J'))
        {
            obj.insertCellByRightShift();
            obj.printGrid();
            obj.printData();
        }
        // if (GetAsyncKeyState('K'))
        // {
        //     obj.insertCellByLeftShift();
        //     obj.printGrid();
        //     obj.printData();
        // }
        if (GetAsyncKeyState('K'))
        {
            obj.insertCellByDownShift();
            obj.printGrid();
            obj.printData();
        }
        // if (GetAsyncKeyState('O'))
        // {
        //     obj.insertCellByUpShift();
        //     obj.printGrid();
        //     obj.printData();
        // }
        if(GetAsyncKeyState(VK_LSHIFT))
        {
            obj.deleteCellByLeftShift();
            obj.printGrid();
            obj.printData();
        }
        if(GetAsyncKeyState(VK_RSHIFT))
        {
            obj.deleteCellByUpShift();
            obj.printGrid();
            obj.printData();
        }
        if ((GetAsyncKeyState(VK_DELETE)) && (GetAsyncKeyState('C')))
        {
            obj.deleteColumn();
            system("cls");
            obj.printGrid();
            obj.printData();
        }
        if ((GetAsyncKeyState(VK_DELETE)) && (GetAsyncKeyState('R')))
        {
            obj.deleteRow();
            system("cls");
            obj.printGrid();
            obj.printData();
        }
        if ((GetAsyncKeyState(VK_HOME)) && (GetAsyncKeyState('C')))
        {
            obj.clearColumn();
            obj.printGrid();
            obj.printData();
        }
        if ((GetAsyncKeyState(VK_HOME)) && (GetAsyncKeyState('R')))
        {
            obj.clearRow();
            obj.printGrid();
            obj.printData();
        }
        if ((c == 'T' || c == 't'))
        {
            // obj.selectMovement();
            obj.selectMovement();
            obj.printGrid();
        }
        if (GetAsyncKeyState(VK_ESCAPE))
        {
            break;
        }
        if (GetAsyncKeyState('V') )
        {
            system("cls");
            obj.paste();
            obj.printGrid();
            obj.printData();
        }
        if (GetAsyncKeyState('Y') )
        {
            obj.saveFile();
        }
        if (GetAsyncKeyState('L') )
        {
            system("cls");
            obj.loadFile();
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