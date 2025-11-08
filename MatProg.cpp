#include <iostream>
#include <vector>
#include <string>
#include <fstream>
#include <sstream>
#include "libxl.h"

using namespace libxl;
using namespace std;

struct SimplexData {
    vector<vector<double>> tableau;
    vector<int> basis;
    vector<string> variables;
    int numVariables;
    int numConstraints;
    bool isMaximization;
    vector<string> steps;
};

// Функции для работы с данными
void initializeSimplexData(SimplexData& data);
bool loadProblemFromFile(SimplexData& data, const string& filename);
void inputProblemFromKeyboard(SimplexData& data);
void convertToCanonicalForm(SimplexData& data);
bool solveSimplexProblem(SimplexData& data);
void printCurrentTableau(const SimplexData& data, int iteration);
void generateExcelReport(const SimplexData& data, const string& filename);
void displayMenu();

// Вспомогательные функции
int findPivotColumn(const SimplexData& data);
int findPivotRow(const SimplexData& data, int pivotCol);
void performPivotOperation(SimplexData& data, int pivotRow, int pivotCol);
bool checkOptimalSolution(const SimplexData& data);
bool checkUnboundedSolution(const SimplexData& data, int pivotCol);

int main() {
    setlocale(LC_ALL, "Russian");
    SimplexData problemData;
    initializeSimplexData(problemData);

    int choice;
    bool problemLoaded = false;

    do {
        displayMenu();
        cout << "Выберите пункт меню: ";
        cin >> choice;

        switch (choice) {
        case 1:
            inputProblemFromKeyboard(problemData);
            problemLoaded = true;
            break;

        case 2: {
            string filename;
            cout << "Введите имя файла: ";
            cin >> filename;
            if (loadProblemFromFile(problemData, filename)) {
                problemLoaded = true;
                cout << "Задача успешно загружена!" << endl;
            }
            else {
                cout << "Ошибка загрузки файла!" << endl;
            }
            break;
        }

        case 3:
            if (problemLoaded) {
                convertToCanonicalForm(problemData);
                cout << "Задача приведена к канонической форме" << endl;
            }
            else {
                cout << "Сначала загрузите задачу!" << endl;
            }
            break;

        case 4:
            if (problemLoaded) {
                if (solveSimplexProblem(problemData)) {
                    cout << "Решение найдено успешно!" << endl;
                }
                else {
                    cout << "Решение не найдено!" << endl;
                }
            }
            else {
                cout << "Сначала загрузите задачу!" << endl;
            }
            break;

        case 5:
            if (problemLoaded) {
                string filename;
                cout << "Введите имя файла для отчета (без расширения): ";
                cin >> filename;
                filename += ".xlsx";
                generateExcelReport(problemData, filename);
                cout << "Отчет сохранен в файл: " << filename << endl;
            }
            else {
                cout << "Сначала загрузите задачу!" << endl;
            }
            break;

        case 6:
            if (problemLoaded) {
                cout << "Текущее состояние задачи:" << endl;
                printCurrentTableau(problemData, 0);
            }
            else {
                cout << "Задача не загружена!" << endl;
            }
            break;

        case 0:
            cout << "Выход из программы..." << endl;
            break;

        default:
            cout << "Неверный выбор! Попробуйте снова." << endl;
        }

        cout << endl;

    } while (choice != 0);

    return 0;
}

void displayMenu() {
    cout << "=== СИМПЛЕКС-МЕТОД РЕШЕНИЯ ЗАДАЧ ЛП ===" << endl;
    cout << "1. Ввод задачи с клавиатуры" << endl;
    cout << "2. Загрузка задачи из файла" << endl;
    cout << "3. Приведение к канонической форме" << endl;
    cout << "4. Решение задачи симплекс-методом" << endl;
    cout << "5. Генерация отчета в Excel" << endl;
    cout << "6. Просмотр текущей симплекс-таблицы" << endl;
    cout << "0. Выход" << endl;
    cout << "----------------------------------------" << endl;
}

void initializeSimplexData(SimplexData& data) {
    data.numVariables = 0;
    data.numConstraints = 0;
    data.isMaximization = true;
    data.tableau.clear();
    data.basis.clear();
    data.variables.clear();
    data.steps.clear();
}

//file reading
bool loadProblemFromFile(SimplexData& data, const string& filename) {
    ifstream file(filename);
    if (!file.is_open()) {
        return false;
    }

    initializeSimplexData(data);
    

    file.close();
    return true;
}

void inputProblemFromKeyboard(SimplexData& data) {
    cout << "Введите количество переменных: ";
    cin >> data.numVariables;

    cout << "Введите количество ограничений: ";
    cin >> data.numConstraints;

    char problemType;
    cout << "Тип задачи (m - максимизация, n - минимизация): ";
    cin >> problemType;
    data.isMaximization = (problemType == 'm' || problemType == 'M');

    cout << "Введите коэффициенты целевой функции:" << endl;
    // Реализация ввода коэффициентов

    cout << "Введите ограничения:" << endl;
    // Реализация ввода ограничений

    initializeSimplexData(data);
}

void convertToCanonicalForm(SimplexData& data) {
    // Реализация приведения к канонической форме
    // Добавление дополнительных переменных
    // Преобразование неравенств в равенства
}

bool solveSimplexProblem(SimplexData& data) {
    int iteration = 0;

    while (!checkOptimalSolution(data)) {
        printCurrentTableau(data, iteration);

        int pivotCol = findPivotColumn(data);
        if (pivotCol == -1) {
            return false; // Решение не найдено
        }

        if (checkUnboundedSolution(data, pivotCol)) {
            cout << "Целевая функция неограничена!" << endl;
            return false;
        }

        int pivotRow = findPivotRow(data, pivotCol);
        if (pivotRow == -1) {
            return false;
        }

        performPivotOperation(data, pivotRow, pivotCol);
        iteration++;
    }

    printCurrentTableau(data, iteration);
    return true;
}

void printCurrentTableau(const SimplexData& data, int iteration) {
    cout << "Симплекс-таблица (итерация " << iteration << "):" << endl;
    // Реализация вывода таблицы в консоль
}

void generateExcelReport(const SimplexData& data, const string& filename) {
    Book* book = xlCreateBook();

    if (book) {
        // Конвертируем имя файла в широкую строку
        wstring wfilename(filename.begin(), filename.end());

        // Создание титульного листа
        Sheet* titleSheet = book->addSheet(L"Титульный лист");
        if (titleSheet) {
            titleSheet->writeStr(1, 0, L"Постановка задачи");
            titleSheet->writeStr(3, 0, L"Исходные данные");
            titleSheet->writeStr(5, 0, L"Тип задачи");

            wstring problemType = data.isMaximization ? L"Максимизация" : L"Минимизация";
            titleSheet->writeStr(5, 1, problemType.c_str());

            // Заголовок
            titleSheet->writeStr(0, 0, L"Решение задачи линейного программирования");
            titleSheet->writeStr(0, 1, L"симплекс-методом");

            // Добавляем информацию о переменных и ограничениях
            titleSheet->writeStr(7, 0, L"Количество переменных:");
            titleSheet->writeNum(7, 1, data.numVariables);

            titleSheet->writeStr(8, 0, L"Количество ограничений:");
            titleSheet->writeNum(8, 1, data.numConstraints);
        }

        // Создание листа с процессом решения
        Sheet* processSheet = book->addSheet(L"Процесс решения");
        if (processSheet) {
            processSheet->writeStr(0, 0, L"Начальная симплекс-таблица");
            processSheet->writeStr(2, 0, L"Промежуточные итерации с выделением разрешающих элементов");

            // Заголовки для симплекс-таблицы
            int col = 0;
            processSheet->writeStr(4, col++, L"Базис");
            for (int i = 0; i < data.numVariables; ++i) {
                wstring varName = L"x" + to_wstring(i + 1);
                processSheet->writeStr(4, col++, varName.c_str());
            }
            processSheet->writeStr(4, col, L"Решение");
        }

        // Создание листа с результатами
        Sheet* resultsSheet = book->addSheet(L"Результаты");
        if (resultsSheet) {
            resultsSheet->writeStr(0, 0, L"Оптимальное значение целевой функции");
            resultsSheet->writeStr(2, 0, L"Значения переменных в оптимальной точке");
            resultsSheet->writeStr(4, 0, L"Статус решения:");
            resultsSheet->writeStr(4, 1, L"Решение найдено");
        }

        // Сохраняем файл с конвертированным именем
        if (book->save(wfilename.c_str())) {
            cout << "Excel файл успешно создан: " << filename << endl;
        }
        else {
            cout << "Ошибка при сохранении Excel файла!" << endl;
        }

        book->release();
    }
    else {
        cout << "Ошибка при создании Excel книги!" << endl;
    }
}

// Реализации вспомогательных функций
int findPivotColumn(const SimplexData& data) {
    // Поиск разрешающего столбца
    return -1;
}

int findPivotRow(const SimplexData& data, int pivotCol) {
    // Поиск разрешающей строки
    return -1;
}

void performPivotOperation(SimplexData& data, int pivotRow, int pivotCol) {
    // Выполнение операции поворота
}

bool checkOptimalSolution(const SimplexData& data) {
    // Проверка оптимальности решения
    return false;
}

bool checkUnboundedSolution(const SimplexData& data, int pivotCol) {
    // Проверка на неограниченность
    return false;
}