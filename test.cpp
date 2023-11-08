#include <iostream>
#include <windows.h>
using namespace std;

// Your Cell struct remains unchanged
template <typename T>
struct Cell
{
    T value;
    Cell* left;
    Cell* right;
    Cell* up;
    Cell* down;
    Cell(T value)
    {
        this->value = value;
        left = nullptr;
        right = nullptr;
        up = nullptr;
        down = nullptr;
    }
    Cell()
    {

    }
};

template <typename T>
class MiniExcel
{
private:
    Cell<T>* current;
    int rows;
    int cols;
    Cell<T>* grid; // 2D array of cells

public:
    MiniExcel()
    {
        current = nullptr;
        rows = 5;
        cols = 5;

        // Create the grid of cells
        grid = new Cell<T>[rows * cols];
        for (int i = 0; i < rows; ++i)
        {
            for (int j = 0; j < cols; ++j)
            {
                Cell<T>* cell = &grid[i * cols + j];
                if (i > 0) cell->up = &grid[(i - 1) * cols + j];
                if (i < rows - 1) cell->down = &grid[(i + 1) * cols + j];
                if (j > 0) cell->left = &grid[i * cols + (j - 1)];
                if (j < cols - 1) cell->right = &grid[i * cols + (j + 1)];
            }
        }

        // Set the initial current cell
        current = &grid[0];
    }

    void printGrid()
    {
        for (int i = 0; i < rows; ++i)
        {
            for (int j = 0; j < cols; ++j)
            {
                std::cout << "+---";
            }
            std::cout << "+\n";

            for (int j = 0; j < cols; ++j)
            {
                std::cout << "|";
                if (&grid[i * cols + j] == current)
                {
                    std::cout << " * ";
                }
                else
                {
                    std::cout << "   ";
                }
            }
            std::cout << "|\n";
        }

        for (int j = 0; j < cols; ++j)
        {
            std::cout << "+---";
        }
        std::cout << "+\n";
    }

    void moveCursorUp()
    {
        if (current->up != nullptr)
        {
            current = current->up;
        }
    }

    void moveCursorDown()
    {
        if (current->down != nullptr)
        {
            current = current->down;
        }
    }

    void moveCursorLeft()
    {
        if (current->left != nullptr)
        {
            current = current->left;
        }
    }

    void moveCursorRight()
    {
        if (current->right != nullptr)
        {
            current = current->right;
        }
    }

    void insertValue(T value)
    {
        current->value = value;
    }

    ~MiniExcel()
    {
        delete[] grid;
    }
};

int main()
{
    MiniExcel<int> excel;
    excel.printGrid();
    while (true)
    {
        cout << "hello";
        if (GetAsyncKeyState(VK_UP))
        {
            excel.moveCursorUp();
        }
        if (GetAsyncKeyState(VK_DOWN))
        {
            excel.moveCursorDown();
        }
        if (GetAsyncKeyState(VK_LEFT))
        {
            excel.moveCursorLeft();
        } 
        if (GetAsyncKeyState(VK_RIGHT))
        {
            excel.moveCursorRight();
        }
        Sleep(100);
        
    }
    // excel.moveCursorRight();
    // excel.moveCursorDown();
    // excel.insertValue(42);
    // excel.printGrid();

    return 0;
}
