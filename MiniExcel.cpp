#include <windows.h>
#include <iostream>
#include <conio.h>
#include <climits>
#include <string>
#include <vector>
#include <fstream>
#pragma once
using namespace std;
struct Cordinate
{
    int row;
    int col;
};
struct Cell
{
    string data;
    Cell* left;
    Cell* right;
    Cell* up;
    Cell* down;

    Cell(string data = "     ")
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
    Cell* rangeStart;
    Cell* rangeEnd;
    vector<vector<string>> clipboard;
    int rows;
    int cols;
    int currentRow;
    int currentCol;
    int cellWidth = 10;
    int cellHeight = 5;
    Cordinate RangeStart{};
    Cordinate RangeEnd{};

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
        // goToRowCol((row * cellHeight) + cellHeight / 2, (col * cellWidth) + cellWidth / 2);
        // cout << "     ";
    }
    void printRow()
    {
       for (int ci = 0, ri = rows + 1; ci < cols; ci++)
		{
			printCell(ri, ci, 7);
		}
        rows++;
        // printGrid();
    }
    void printData()
    {
        Cell* temp = head;
        for (int i = 0; i < rows; i++)
        {
            Cell* temp2 = temp;
            for (int j = 0; j < cols; j++)
            {
                printCellData(i, j, temp, 10);
                temp = temp -> right;
            }
            temp = temp2 -> down;
        }
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
    Cell* newCol()
    {
        Cell* temp = new Cell();
        Cell* curr = temp;
        for (int i = 0; i < rows - 1; i++)
        {
            Cell* temp2 = new Cell();
            temp -> down = temp2;
            temp2 -> up = temp;
            temp = temp2;
        }
        return curr;
    }
    void insertRowBelow()
    {
        Cell* temp2 = newRow();
        Cell* temp1 = current;
        // loop to reach the start of current list
        while (temp1 -> left != nullptr)
        {
            temp1 = temp1 -> left;
        }
        if (temp1 -> down == nullptr)
        {
            while (temp1 != nullptr)
            {
                temp1 -> down = temp2;
                temp2 -> up = temp1;
                temp1 = temp1 -> right;
                temp2 = temp2 -> right;
            }
        }
        else 
        {
            Cell* temp3 = temp1 -> down;
            while (temp1 != nullptr)
            {
                temp3 -> up = temp2;
                temp1 -> down = temp2;
                temp2 -> up = temp1;
                temp2 -> down = temp3;
                temp1 = temp1 -> right;
                temp2 = temp2 -> right;
                temp3 = temp3 -> right;
            }
        }
        rows++;
    }
    void insertRowAbove()
    {
        Cell* temp2 = newRow();
        Cell* temp1 = current;
        // loop to reach the start of current list
        while (temp1 -> left != nullptr)
        {
            temp1 = temp1 -> left;
        }
        if (temp1 -> up == nullptr)
        {
            head = temp2;
            while (temp1 != nullptr)
            {
                temp1 -> up = temp2;
                temp2 -> down = temp1;
                temp1 = temp1 -> right;
                temp2 = temp2 -> right;
            }
        }
        else
        {
           Cell* temp3 = temp1 -> up;
            while (temp1 != nullptr)
            {
                temp1 -> up = temp2;
                temp2 -> up = temp3;
                temp3 -> down = temp2;
                temp2 -> down = temp1;
                temp1 = temp1 -> right;
                temp2 = temp2 -> right;
                temp3 = temp3 -> right;
            }
        }
        currentRow++;
        rows++;
        // printGrid();
        // printData();
    }
    void insertColumnToRight()
    {
        Cell* temp2 = newCol();
        Cell* temp1 = current;
        while(temp1 -> up != nullptr)
        {
            temp1 = temp1 -> up;
        }
        if (temp1 -> right == nullptr)
        {
            while (temp1 != nullptr)
            {
                temp1 -> right = temp2;
                temp2 -> left = temp1;
                temp1 = temp1 -> down;
                temp2 = temp2 -> down;
            }
        }
        else 
        {
            Cell* temp3 = temp1 -> right;
            while (temp1 != nullptr)
            {
                temp1 -> right = temp2;
                temp2 -> right = temp3;
                temp3 -> left = temp2;
                temp2 -> left = temp1;
                temp1 = temp1 -> down;
                temp2 = temp2 -> down;
                temp3 = temp3 -> down;
            }
        }
        cols++;
    }
    void insertColumnToLeft()
    {
        Cell* temp1 = current;
        Cell* temp2 = newCol();
        while(temp1 -> up != nullptr)
        {
            temp1 = temp1 -> up;
        }
        if (temp1 -> left == nullptr)
        {
            while (temp1 != nullptr)
            {
                temp1 -> left = temp2;
                temp2 -> right = temp1;
                temp1 = temp1 -> down;
                temp2 = temp2 -> down;
            }
            head = head -> left;
        }
        else
        {
            Cell* temp3 = temp1 -> left;
            while (temp1 != nullptr)
            {
                temp1 -> left = temp2;
                temp2 -> right = temp1;
                temp3 -> right = temp2;
                temp2 -> left = temp3;
                temp1 = temp1 -> down;
                temp2 = temp2 -> down;
                temp3 = temp3 -> down;
            }
        }
        cols++;
        currentCol++;
    }
    void insertCellByRightShift()
    {
        Cell* temp = current;
        while(current -> right != nullptr)
        {
            current = current -> right;
        }
        insertColumnToRight();
        current = current -> right;
        while (current != temp)
        {
            current -> data = current -> left -> data;
            current = current -> left;
        }
        current -> data = "    ";
    }
    void insertCellByLeftShift()
    {
        Cell* temp = current;
        while (current -> left != nullptr)
        {
            current = current -> left;
        }
        insertColumnToLeft();
        current = current -> left;
        while (current != temp)
        {
            current -> data = current -> right -> data;
            current = current -> right;
        }
        current -> data = "    ";
    }
    void insertCellByDownShift()
    {
        Cell* temp = current;
        while (current -> down != nullptr)
        {
            current = current -> down;
        }
        insertRowBelow();
        current = current -> down;
        while (current != temp)
        {
            current -> data = current -> up -> data;
            current = current -> up;
        }
        current -> data = "     ";
    }
    void insertCellByUpShift()
    {
        Cell* temp = current;
        while (current -> up != nullptr)
        {
            current = current -> up;
        }
        insertRowAbove();
        current = current -> up;
        while (current != temp)
        {
            current -> data = current -> down -> data;
            current = current -> down;
        }
        current -> data = "     ";
    }
    void deleteCellByLeftShift()
    {
        Cell* temp = current;
        while (temp -> right != nullptr)
        {
            temp -> data = temp -> right -> data;
            temp = temp -> right;
        }
        if (current -> left != nullptr)
        {
            current = current -> left;
        }
        currentCol--;
    }
    void deleteCellByUpShift()
    {
        Cell* temp = current;
        while (temp -> down != nullptr)
        {
            temp -> data = temp -> down -> data;
            temp = temp -> down;
        }
        if (current -> up != nullptr)
        {
            current = current -> up;
        }
        currentRow--;
    }
    void deleteColumn()
    {
        if (cols <= 1)
        {
            return;
        }
        Cell* temp1 = current;
        while(temp1 -> up != nullptr)
        {
            temp1 = temp1 -> up;
        }
        if (temp1 -> left == nullptr)
        {
            temp1 = temp1 -> right;
            current = current -> right;
            head = head -> right;
            while (temp1 != nullptr)
            {
                temp1 -> left = nullptr;
                temp1 = temp1 -> down;
            }
        }
        else if (temp1 -> right == nullptr)
        {
            temp1 = temp1 -> left;
            current = current -> left;
            while (temp1 != nullptr)
            {
                temp1 -> right = nullptr;
                temp1 = temp1 -> down;
            }
            currentCol--;
        }
        else
        {
            Cell* temp2 = temp1 -> right;
            temp1 = temp1 -> left;
            // Cell* temp3 = temp1;
            current = current -> left;
            while(temp1 != nullptr)
            {
                temp1 -> right = temp2;
                temp2 -> left = temp1;
                temp1 = temp1 -> down;
                temp2 = temp2 -> down;
            }
            currentCol--;
        }
        cols--;
    }
    void deleteRow()
    {
        if (rows <= 1)
        {
            return;
        }
        Cell* temp1 = current;
        while (temp1 -> left != nullptr)
        {
            temp1 = temp1 -> left;
        }
        if (temp1 -> up == nullptr)
        {
            current = current -> down;
            head = head -> down;
            temp1 = temp1 -> down;
            while (temp1 != nullptr)
            {
                temp1 -> up = nullptr;
                temp1 = temp1 -> right;
            }
        }
        else if (temp1 -> down == nullptr)
        {
            current = current -> up;
            temp1 = temp1 -> up;
            while (temp1 != nullptr)
            {
                temp1 -> down = nullptr;
                temp1 = temp1 -> right;
            }
            currentRow--;
        }
        else 
        {
            Cell* temp2 = temp1 -> down;
            temp1 = temp1 -> up;
            current = current -> up;
            while(temp1 != nullptr)
            {
                temp2 -> up = temp1;
                temp1 -> down = temp2;
                temp1 = temp1 -> right;
                temp2 = temp2 -> right;
            }
            currentRow--;
        }
        rows--;
    }
    void clearColumn()
    {
        Cell* temp1 = current;
        while (temp1 -> up != nullptr)
        {
            temp1 = temp1 -> up;
        }
        while (temp1 -> down != nullptr)
        {
            temp1 -> data = "     ";
            temp1 = temp1 -> down;
        }
    }
    void clearRow()
    {
        Cell* temp1 = current;
        while (temp1 -> left != nullptr)
        {
            temp1 = temp1 -> left;
        }
        while(temp1 -> right != nullptr)
        {
            temp1 -> data = "     ";
            temp1 = temp1 -> right;
        }
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
    void insertValue(string value)
    {
        bool flag = isDigit(value);
        if (flag)
        {
            current -> data = value;
        }
        else 
        {
            return;
        }
        eraseCell(currentRow, currentCol, 10);
        printCellData(currentRow, currentCol, current , 10);
    }
    bool isDigit(string value)
    {
        if (value.empty())
        {
            return false;
        }
        int len = value.length();
        for (int i = 0; i < len; i++)
        {
            if(!isdigit(value[i]))
            {
                return false;
            }
        }
        return true;
    }
    void printCellData(int row, int col, Cell* data, int k)
    {
        colour(k);
        goToRowCol((row * cellHeight) + cellHeight / 2, (col * cellWidth) + cellWidth / 2);
        cout << data -> data;
    }
    void eraseCell(int row, int col, int k)
    {
        colour(k);
        goToRowCol((row * cellHeight) + cellHeight / 2, (col * cellWidth) + cellWidth / 2);
        cout << "     ";
    }
    void selectMovement()
    {
        rangeStart = current;
        RangeStart.row = currentRow;
        RangeStart.col = currentCol;
        printCell(currentRow, currentCol, 12);
        while (true)
        {
            char c = _getch();
            if (c == 'w' || c == 'W') // up w
            {
                moveUp();
            }
            else if (c == 's' || c == 'S') // down s
            {
                moveDown();
            }
            else if (c == 'a' || c == 'A') // left a
            {
                moveLeft();
            }
            else if (c == 'd' || c == 'D') // right d
            {
               moveRight();
            }
            else if ((c == 'E' || c == 'e'))
            {
                printCell(currentRow, currentCol, 12);
                rangeEnd = current;
                break;
            }

        }
        RangeEnd.col = currentCol;
        RangeEnd.row = currentRow; 
        while (true)
        {
            char c = _getch();
            if (c == 'w' || c == 'W') // up w
            {
                moveUp();
            }
            else if (c == 's' || c == 'S') // down s
            {
                moveDown();
            }
            else if (c == 'a' || c == 'A') // left a
            {
                moveLeft();
            }
            else if (c == 'd' || c == 'D') // right d
            {
               moveRight();
            }
            else if ((GetAsyncKeyState('O')) )
            {
                int count = GetRangeCount();
                current -> data = to_string(count);
                printCellData(currentRow, currentCol, current, 12);
                break;
            }
            else if (GetAsyncKeyState('Z'))
            {
                int sum = GetRangeSum();
                int count = GetRangeCount();
                current -> data = to_string(sum / count);
                printCellData(currentRow, currentCol, current, 12);
                break;
            }
            else if (GetAsyncKeyState('U') )
            {
                int sum = GetRangeSum();
                current -> data = to_string(sum);
                printCellData(currentRow, currentCol, current, 12);
                break;
            }
             else if (GetAsyncKeyState('M') )
            {
                int max = getRangeMaximum();
                current -> data = to_string(max);
                printCellData(currentRow, currentCol, current, 12);
                break;
            }
             else if (GetAsyncKeyState('N') )
            {
                int min = getRangeMinimum();
                current -> data = to_string(min);
                printCellData(currentRow, currentCol, current, 12);
                break;
            }
            else if (GetAsyncKeyState('C') )
            {
                copy();
                break;
            }
            else if (GetAsyncKeyState('X') )
            {
                cut();
                printData();
                break;
            }
           
        }
       
    }
    int GetRangeSum()
    {
        int sum = 0;
        Cell* temp = rangeStart;
        for (int i = RangeStart.row; i <= RangeEnd.row; i++)
        {
            Cell* temp2 = temp;
            for (int j = RangeStart.col; j <= RangeEnd.col; j++)
            {
                if (isDigit(temp -> data))
                {
                    sum += stoi(temp -> data);
                }
                temp = temp -> right;
            }
            temp = temp2 -> down;
        }
        return sum;
    }
    int GetRangeCount()
    {
        int count = 0;
        Cell* temp = rangeStart;
        for (int i = RangeStart.row; i <= RangeEnd.row; i++)
        {
            Cell* temp2 = temp;
            for (int j = RangeStart.col; j <= RangeEnd.col; j++)
            {
                if (!isDigit(temp -> data))
                {
                    count++;
                }
                temp = temp -> right;
            }
            temp = temp2 -> down;
        }
        return count;
    }
    int getRangeMinimum()
    {
        int min = INT_MAX;
        Cell* temp = rangeStart;
        for (int i = RangeStart.row; i <= RangeEnd.row; i++)
        {
            Cell* temp2 = temp;
            for (int j = RangeStart.col; j <= RangeEnd.col; j++)
            {
                if (isDigit(temp -> data))
                {
                   int t = stoi(temp -> data);
                   if (t < min)
                   {
                    min = t;
                   }
                }
                temp = temp -> right;
            }
            temp = temp2 -> down;
        }
        return min;
    }
    int getRangeMaximum()
    {
        int max = INT_MIN;
        Cell* temp = rangeStart;
        for (int i = RangeStart.row; i <= RangeEnd.row; i++)
        {
            Cell* temp2 = temp;
            for (int j = RangeStart.col; j <= RangeEnd.col; j++)
            {
                if (isDigit(temp -> data))
                {
                   int t = stoi(temp -> data);
                   if (t > max )
                   {
                    max = t;
                   }
                }
                temp = temp -> right;
            }
            temp = temp2 -> down;
        }
        return max;
    }
    void copy()
    {
        clipboard.clear();
        Cell* temp = rangeStart;
        for (int i = RangeStart.row; i <= RangeEnd.row; i++)
        {
            vector<string> clip;
            Cell* temp2 = temp;
            for (int j = RangeStart.col; j <= RangeEnd.col; j++)
            {
                clip.push_back(temp -> data);
                temp = temp -> right;
            }
            clipboard.push_back(clip);
            temp = temp2 -> down;
        }
    }
    void cut()
    {
        clipboard.clear();
        Cell* temp = rangeStart;
        for (int i = RangeStart.row; i <= RangeEnd.row; i++)
        {
            vector<string> clip;
            Cell* temp2 = temp;
            for (int j = RangeStart.col; j <= RangeEnd.col; j++)
            {
                clip.push_back(temp -> data);
                temp -> data = "     ";
                temp = temp -> right;
            }
            clipboard.push_back(clip);
            temp = temp2 -> down;
        }
    }
    void paste()
    {
        Cell* temp = current;
        for (int i = 0; i < clipboard.size(); i++)
        {
            Cell* temp2 = current;
            for (int j = 0; j < clipboard[0].size(); j++)
            {
                current -> data = clipboard[i][j];
                if (current -> right == nullptr)
                {
                    insertColumnToRight();
                }
                current = current -> right;
            }
            if (temp2 -> down == nullptr)
            {
                insertRowBelow();
            }
            current = temp2 -> down;
        }
        current = temp;
    }
    void saveFile()
    {
        Cell* temp = head;
        ofstream file("save.txt");
        file << rows << endl;
        file << cols << endl;
        for (int i = 0; i < rows; i++)
        {
            Cell* temp2 = temp;
            for (int j = 0; j < cols; j++)
            {
                if (temp -> data == "     ")
                {
                    file << "space" << " ";
                }
                else
                {
                    file << temp -> data << " ";
                }
                temp = temp -> right;
            }
            file << endl;
            temp = temp2 -> down;
        }
    }
    void loadFile()
    {
        ifstream file("save.txt");
        file >> rows;
        file >> cols;
        current = nullptr;
        head = nullptr;
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
        string data;
        Cell* temp = current;
        for (int i = 0; i < rows; i++)
        {
            Cell* temp2 = temp;
            for (int j = 0; j < cols; j++)
            {
                file >> data;
                if (data == "space")
                {
                    temp -> data = "     ";
                }
                else
                {
                    temp -> data = data;
                }
                temp = temp -> right;
            }
            temp = temp2 -> down;
        }
    }
    void performOperation()
    {
        rangeStart = current;
        int sum = 0;
        int average = 0;
        int count = 0; 
        int max = INT_MIN;
        int min = INT_MAX;
        while (true)
        {
            char c = _getch();
            string r = current -> data;
            string numeric = "";
            for (char c : r)
            {
                if (isdigit(c))
                {
                    numeric += c;
                }
            }
            if (!numeric.empty())
            {
                int temp = std::stoi(r);
                if (temp < min)
                {
                    min = temp;
                }
                if (temp > max)
                {
                    max = temp;
                }
                sum += temp;
                count++;
            }
            if (c == 'w' || c == 'W') // up w
            {
                if (current -> up != nullptr)
                {
                    current = current -> up;
                    currentRow--;
                }
            }
            else if (c == 's' || c == 'S') // down s
            {
                if (current -> down != nullptr)
                {
                    current = current -> down;
                    currentRow++;
                }
            }
            else if (c == 'a' || c == 'A') // left a
            {
                if (current -> left != nullptr)
                {
                    current = current -> left;
                    currentCol--;
                }
            }
            else if (c == 'd' || c == 'D') // right d
            {
                if (current -> right != nullptr)
                {
                    current = current -> right;
                    currentCol++;
                }
            }
            else if ((c == 'E' || c == 'e'))
            {
                break;
            }
            printCell(currentRow, currentCol, 10);
        }
        char c = _getch();
        if (c == 's' || c == 'S')
        {
            current -> data = to_string(sum);
            printCellData(currentRow,currentCol, current, 12);
        }
        else if (c == 'a' || c == 'A')
        {
            average = sum / count;
            current -> data = to_string(average);
            printCellData(currentRow, currentCol, current, 12);
        }
        else if (c == 'c' || c == 'C')
        {
            current -> data = to_string(count);
            printCellData(currentRow, currentCol, current, 12);
        }
        else if (c == 'x' || c == 'X')
        {
            current -> data = to_string(max);
            printCellData(currentRow, currentCol, current, 12);
        }
        else if (c == 'n' || c == 'N')
        {
            current -> data = to_string(min);
            printCellData(currentRow, currentCol, current, 12);
        }
    }
};