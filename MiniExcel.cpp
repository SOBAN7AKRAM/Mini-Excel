#include <windows.h>
#include <iostream>
#include <conio.h>
#pragma once
using namespace std;
struct Cell
{
    string data;
    Cell* left;
    Cell* right;
    Cell* up;
    Cell* down;
    Cell(string data = " ")
    {
        this->data = data;
        left = nullptr;
        right = nullptr;
        up = nullptr;
        down = nullptr;
    }
};
class MiniExcel
{
private:
    Cell* current;
    Cell* head;
    int rows;
    int cols;
    int currentRow;
    int currentCol;
    int cellWidth = 10;
    int cellHeight = 5;
public:
    MiniExcel()
    {
        current = nullptr;
        head = nullptr;
        this -> rows = 5;
        this -> cols = 5;
        currentRow = currentCol = 0;
        head = newRow();
        current = head;
        for (int i = 0; i < rows - 1; i++)
        {
            insertRowBelow();
            rows--; // rows-- because insertRowBelow function incrementing the value of rows
            current = current -> down;
        }
        current = head;
        printGrid();
    }
    void colour(int k)
    {
        if (k > 255)
        {
            k = 1;
        }
        HANDLE hOutput = GetStdHandle(STD_OUTPUT_HANDLE);
        SetConsoleTextAttribute(hOutput, k);
    }
    void goToRowCol(int rpos, int cpos)
    {
        COORD scrn;
        HANDLE hOutput = GetStdHandle(STD_OUTPUT_HANDLE);
        scrn.X = cpos;
        scrn.Y = rpos;
        SetConsoleCursorPosition(hOutput, scrn);
    }
    void printCell(int row, int col, int k)
    {
        char c = 219;
        colour(k);
        goToRowCol(row * cellHeight, col * cellWidth);
        for (int i = 0; i < cellWidth; i++)
        {
            cout << c;
        }
        goToRowCol(row * cellHeight + cellHeight, col * cellWidth);
        for (int i = 0; i <= cellWidth; i++)
        {
            cout << c;
        }
        for (int i = 0; i < cellHeight; i++)
        {
            goToRowCol(row * cellHeight + i, col * cellWidth);
            cout << c;
        }
        for (int i = 0; i <= cellHeight; i++)
        {
            goToRowCol(row * cellHeight + i, col * cellWidth + cellWidth);
            cout << c;
        }
        goToRowCol((row * cellHeight) + cellHeight / 2, (col * cellWidth) + cellWidth / 2);
        cout << "     ";
    }
    void printGrid()
    {
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                printCell(i, j, 7);
            }
        }
    }
    Cell* newRow()
    {
        Cell* temp = new Cell();
        Cell* curr = temp;
        for (int i = 0; i < cols - 1; i++) // cols - 1 because we created a Cell before loop
        {
            Cell* temp2 = new Cell();
            temp -> right = temp2;
            temp2 -> left = temp;
            temp = temp2; 
        }
        return curr;
    }
    Cell* insertRowBelow()
    {
        Cell* temp2 = newRow();
        Cell* temp1 = current;
        // loop to reach the start of current list
        while (temp1 -> left != nullptr)
        {
            temp1 = temp1 -> left;
        }
        while (temp1 != nullptr)
        {
            temp1 -> down = temp2;
            temp2 -> up = temp1;
            temp1 = temp1 -> right;
            temp2 = temp2 -> right;
        }
        rows++;
    }
    void moveDown()
    {
        if (current -> down != nullptr)
        {
            printCell(currentRow, currentCol, 7);
            current = current -> down;
            currentRow++;
            printCell(currentRow, currentCol, 10);
        }
    }
    void moveUp()
    {
        if (current -> up != nullptr)
        {
            printCell(currentRow, currentCol, 7);
            current = current -> up;
            currentRow--;
            printCell(currentRow, currentCol, 10);
        }
    }
    void moveLeft()
    {
        if (current -> left != nullptr)
        {
            printCell(currentRow, currentCol, 7);
            current = current -> left;
            currentCol--;
            printCell(currentRow, currentCol, 10);
        }
    }
    void moveRight()
    {
        if (current -> right != nullptr)
        {
            printCell(currentRow, currentCol, 7);
            current = current -> right;
            currentCol++;
            printCell(currentRow, currentCol, 10);
        }
    }


};