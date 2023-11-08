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