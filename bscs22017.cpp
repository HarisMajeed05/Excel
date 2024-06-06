#include <windows.h>
#include <iostream>
#include <conio.h>
#include <iomanip>
#include <string>
#include <fstream>
#include <sstream> 
#include<vector>
using namespace std;

void getRowColbyLeftClick(int& rpos, int& cpos)
{
	HANDLE hInput = GetStdHandle(STD_INPUT_HANDLE);
	DWORD Events;
	INPUT_RECORD InputRecord;
	SetConsoleMode(hInput, ENABLE_PROCESSED_INPUT | ENABLE_MOUSE_INPUT | ENABLE_EXTENDED_FLAGS);
	do
	{
		ReadConsoleInput(hInput, &InputRecord, 1, &Events);
		if (InputRecord.Event.MouseEvent.dwButtonState == FROM_LEFT_1ST_BUTTON_PRESSED)
		{
			cpos = InputRecord.Event.MouseEvent.dwMousePosition.X;
			rpos = InputRecord.Event.MouseEvent.dwMousePosition.Y;
			break;
		}
	} while (true);
}
void gotoRowCol(int rpos, int cpos)
{
	COORD scrn;
	HANDLE hOuput = GetStdHandle(STD_OUTPUT_HANDLE);
	scrn.X = cpos;
	scrn.Y = rpos;
	SetConsoleCursorPosition(hOuput, scrn);
}
void color(int clr)
{
	SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), clr);
}

class myExcel {
	int cellWidth, cellHeight;
	class cell {
		string data;
		cell* right;
		cell* left;
		cell* up;
		cell* down;
		friend class myExcel;
	public:
		cell(string da = "", cell* l = nullptr, cell* r = nullptr, cell* u = nullptr, cell* d = nullptr) {
			data = da;
			left = l;
			right = r;
			up = u;
			down = d;
		}
	};
	int r_size = 0, c_size = 0;
	int r_curr = 0, c_curr = 0;
	cell* curr;
	cell* head;
	cell* tail;
	cell* rangestart = nullptr;
	cell* rangeend = nullptr;
	int cs, rs, re, ce;
	vector<string> v;
public:
	friend class cell;
	myExcel() {
		cellWidth = 7;
		cellHeight = 2;
		r_size = 1;
		c_size = 1;
		head = new cell();
		tail = head;
		curr = head;
		r_curr = 0;
		c_curr = 0;
		v.resize(0);
		for (int i = 1; i < 3; i++) {
			insertColRight();
			curr = curr->right;
		}
		curr = head;
		for (int i = 1; i < 3; i++) {
			insertRowBelow();
			curr = curr->down;
		}
		curr = head;
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
		inputkeyboard();
	}
	void printData() {
		color(7);
		cell* temp = head;
		cell* storCurr = curr;
		int storCurrRow = r_curr;
		int storCurrCol = c_curr;
		r_curr = 0;
		while (temp != nullptr) {
			c_curr = 0;
			curr = temp;
			while (curr != nullptr) {
				printincell();
				curr = curr->right;
				c_curr++;

			}
			cout << endl;
			temp = temp->down;
			r_curr++;
		}
		r_curr = storCurrRow;
		c_curr = storCurrCol;
		curr = storCurr;
	}
	void printincell() {
		color(7);
		int r = r_curr * cellHeight + cellHeight / 2;
		int c = c_curr * cellWidth + cellWidth / 2;
		gotoRowCol(r, c-1);
		if (curr->data == "\0")
			cout << "";
		cout << curr->data;
	}
	void printincell(int row, int col) {
		color(7);
		int r = row * cellHeight + cellHeight / 2;
		int c = col * cellWidth + cellWidth / 2;
		gotoRowCol(r, c-1);
		cout << curr->data;
	}
	void eraseincell() {
		color(7);
		int r = r_curr * cellHeight + cellHeight / 2;
		int c = c_curr * cellWidth + cellWidth / 2;
		gotoRowCol(r, c);
		curr->data = "";
		cout << curr->data;
	}
	void printcellborder(int row, int col, int clr = 7) {
		color(clr);
		int r_curr1 = row;
		int c_curr1 = col;
		cellWidth = 7;
		cellHeight = 2;
		int w = 0, h = 0;
		int r = r_curr1 * cellHeight;
		int c = c_curr1 * cellWidth;
		gotoRowCol(r, c);
		for (w = 0; w < cellWidth; w++) {
			cout << "*";
		}
		w = 0;
		r = (r_curr1 * cellHeight) + cellHeight;
		gotoRowCol(r, c);
		for (w = 0; w <= cellWidth; w++) {
			cout << "*";
		}
		r = r_curr1 * cellHeight;
		c = c_curr1 * cellWidth;
		for (h = 0; h < cellHeight; h++) {
			gotoRowCol(r + h, c);
			cout << "*";
		}
		c = (c_curr1 * cellWidth) + cellWidth;
		r = r_curr1 * cellHeight;
		for (h = 0; h < cellHeight; h++) {
			gotoRowCol(r + h, c);
			cout << "*";
		}
	}
	void printGrid() {
		for (int i = 0; i < r_size; i++) {
			for (int c = 0; c < c_size; c++) {
				printcellborder(i, c);
			}
		}
	}

	void insertColRight() {
		cell* temp = curr;
		while (curr->up != nullptr) {
			curr = curr->up;
		}
		while (curr != nullptr) {
			insertCelRight(curr);
			curr = curr->down;
		}
		curr = temp;
		c_size++;
	}
	cell* insertCelRight(cell*& c) {
		cell* temp = new cell();
		temp->left = c;
		if (c->right != nullptr) {
			temp->right = c->right;
			temp->right->left = temp;
		}
		c->right = temp;
		if (c->up != nullptr && c->up->right != nullptr) {
			temp->up = c->up->right;
			c->up->right->down = temp;
		}
		if (c->down != nullptr && c->down->right != nullptr) {
			temp->down = c->down->right;
			c->down->right->up = temp;
		}
		return temp;
	}
	void insertRowBelow() {
		cell* temp = curr;
		while (curr->left != nullptr) {
			curr = curr->left;
		}
		while (curr != nullptr) {
			insertCellBelow(curr);
			curr = curr->right;
		}
		curr = temp;
		r_size++;
	}
	cell* insertCellBelow(cell*& c) {
		cell* temp = new cell();
		temp->up = c;
		if (c->down != nullptr) {
			temp->down = c->down;
			temp->down->up = temp;
		}
		curr->down = temp;
		if (c->left != nullptr && c->left->down != nullptr) {
			temp->left = c->left->down;
			c->left->down->right = temp;
		}
		if (c->right != nullptr && c->right->down != nullptr) {
			temp->right = c->right->down;
			c->right->down->left = temp;
		}
		return temp;
	}
	cell* insertCellAbove(cell*& c) {
		cell* temp = new cell();
		temp->down = c;
		if (c->up != nullptr) {
			temp->up = c->up;
			temp->up->down = temp;
		}
		c->up = temp;
		if (c->left != nullptr && c->left->up != nullptr) {
			temp->left = c->left->up;
			c->left->up->right = temp;
		}
		if (c->right != nullptr && c->right->up != nullptr) {
			temp->right = c->right->up;
			c->right->up->left = temp;
		}
		return temp;
	}
	void insertRowAbove() {
		cell* temp = curr;
		while (curr->left != nullptr) {
			curr = curr->left;
		}
		while (curr != nullptr) {
			insertCellAbove(curr);
			curr = curr->right;
		}
		curr = temp;
		r_size++;
	}
	cell* insertCellLeft(cell*& c) {
		cell* temp = new cell();
		temp->right = c;
		if (c->left != nullptr) {
			temp->left = c->left;
			temp->left->right = temp;
		}
		c->left = temp;
		if (c->up != nullptr && c->up->left != nullptr) {
			temp->up = c->up->left;
			c->up->left->down = temp;
		}
		if (c->down != nullptr && c->down->left != nullptr) {
			temp->down = c->down->left;
			c->down->left->up = temp;
		}
		return temp;
	}
	void insertColLeft() {
		cell* temp = curr;
		while (curr->up != nullptr) {
			curr = curr->up;
		}
		while (curr != nullptr) {
			insertCellLeft(curr);
			curr = curr->down;
		}
		curr = temp;
		c_size++;
	}

	void InsertCellByRightShift() {
		if (curr == nullptr)return;
		cell* currentCell = curr;
		if (currentCell->right != nullptr) {
			curr = currentCell->right;
			curr->data = currentCell->right->data;
		}
		else {
			insertColRight();
			curr = currentCell->right;
			curr->data = currentCell->right->data;
		}
		cell* newCell = insertCelRight(currentCell);
		curr = newCell;
		if (c_curr == c_size - 1)
			c_size++;
	}
	void InsertCellByDownShift() {
		if (curr == nullptr)return;
		cell* currentCell = curr;
		if (currentCell->down != nullptr)
			curr = currentCell->down;
		else {
			insertRowBelow();
			curr = currentCell->down;
		}
		cell* newCell = insertCellBelow(currentCell);
		curr = newCell;
		if (r_curr == r_size - 1)
			r_size++;
	}



	void clearRow(){
		cell* temp = curr;
		temp->data = curr->data;
		int r = c_curr;
		while (curr->left != nullptr) {
			curr = curr->left;
			c_curr--;
		}
		while (curr->right != nullptr) {
			curr->data = "";
			//printincell(r_curr,c_curr);
			curr = curr->right;
			c_curr++;
		}
		curr = temp;
		curr->data = "";
		c_curr = r;
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}
	void clearCol() {
		cell* temp = curr;
		temp->data = curr->data;
		int r = r_curr;
		while (curr->up != nullptr) {
			curr = curr->up;
			r_curr--;
		}
		while (curr->down != nullptr) {
			curr->data = "";
			curr = curr->down;
			r_curr++;
		}
		curr = temp;
		curr->data = "";
		r_curr = r;
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}
	void delCol() {
		if (c_size <= 1) return;
		if (curr->right != nullptr)
			curr->right->left = curr->left;
		if (curr->left != nullptr)
			curr->left->right = curr->right;
		if (curr->right != nullptr)
			curr = curr->right;
		else
			curr = head;
		c_size--;
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}
	void delRow() {
		if (r_size <= 1) return;
		cell* temp = curr;
		while (curr->left != nullptr) {
			curr = curr->left;
		}
		while (curr != nullptr) {
			if (curr->down != nullptr) {
				cell* downCell = curr->down;
				curr->down = downCell->down;
				if (downCell->down != nullptr) {
					downCell->down->up = curr;
				}
				delete downCell;
			}
			curr = curr->right;
		}
		curr = temp->down;
		if (curr == nullptr) curr = head;
		this->r_size--;
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}

	void inputkeyboard() {
		string d;
		while (1) {
			char k = _getch();
			int ascii = k;
			if (ascii == 27) {  //for ESC
				break;
			}
			else if (ascii == -32) {
				char c = _getch();
				ascii = c;

				if (ascii == 72) {  //up arrow
					if (curr->up == nullptr) {
						insertRowAbove();
						printGrid();
					}
					printcellborder(r_curr, c_curr, 7);
					curr = curr->up;
					r_curr--;
					printcellborder(r_curr, c_curr, 4);
				}
				else if (ascii == 75) {  //left arrow
					if (curr->left == nullptr) {
						insertColLeft();
						printGrid();
					}
					printcellborder(r_curr, c_curr, 7);
					curr = curr->left;
					c_curr--;
					printcellborder(r_curr, c_curr, 4);
				}
				else if (ascii == 77) {  //right arrow
					if (curr->right == nullptr) {
						insertColRight();
						printGrid();
					}
					printcellborder(r_curr, c_curr, 7);
					curr = curr->right;
					c_curr++;
					printcellborder(r_curr, c_curr, 4);
				}
				else if (ascii == 80) {  //down arrow
					if (curr->down == nullptr) {
						insertRowBelow();
						printGrid();
					}
					printcellborder(r_curr, c_curr, 7);
					curr = curr->down;
					r_curr++;
					printcellborder(r_curr, c_curr, 4);
				}
			}
			//else if (ascii == 100) {  //right  == d
			//	if (curr->right == nullptr) {
			//		insertColRight();
			//		printGrid();
			//	}
			//	printcellborder(r_curr, c_curr, 7);
			//	curr = curr->right;
			//	c_curr++;
			//	printcellborder(r_curr, c_curr, 4);
			//}
			//else if (ascii == 97) {  //left == a
			//	if (curr->left == nullptr) {
			//		insertColLeft();
			//		printGrid();
			//	}
			//	printcellborder(r_curr, c_curr, 7);
			//	curr = curr->left;
			//	c_curr--;
			//	printcellborder(r_curr, c_curr, 4);
			//}
			//else if (ascii == 119) {  //up == w
			//	if (curr->up == nullptr) {
			//		insertRowAbove();
			//		printGrid();
			//	}
			//	printcellborder(r_curr, c_curr, 7);
			//	curr = curr->up;
			//	r_curr--;
			//	printcellborder(r_curr, c_curr, 4);
			//}
			//else if (ascii == 115) {  //down == s
			//	if (curr->down == nullptr) {
			//		insertRowBelow();
			//		printGrid();
			//	}
			//	printcellborder(r_curr, c_curr, 7);
			//	curr = curr->down;
			//	r_curr++;
			//	printcellborder(r_curr, c_curr, 4);
			//}
			else if (ascii == 13) { //insert row col by pressing ENTER
				k = _getch();
				ascii = k;
				if (ascii == 100) { //d for right col
					insertColRight();
				}
				if (ascii == 115) { //s for row below
					insertRowBelow();
				}
				if (ascii == 97) { //a for left col
					insertColLeft();
				}
				if (ascii == 119) { //w for row above
					insertRowAbove();
				}
				
				system("cls");
				printGrid();
				printData();
				printButton();
				printcellborder(r_curr, c_curr, 4);
			}
			else if (ascii == 108) {
				InsertCellByRightShift();
				system("cls");
				printGrid();
				printData();
				printButton();
				printcellborder(r_curr, c_curr, 4);
			}
			else if (ascii == 107) {
				InsertCellByDownShift();
				system("cls");
				printGrid();
				printData();
				printButton();
				printcellborder(r_curr, c_curr, 4);
			}
			else if (ascii == 8) { //BACKSPACE for clearing rows/cols
				k = _getch();
				ascii = k;
				if (ascii == 114) { //r for current row
					clearRow();
				}
				if (ascii == 99) { //c for col current
					clearCol();
				}
				//printincell();
				printGrid();
				printData();
				printcellborder(r_curr, c_curr, 4);
			}
			else if (ascii == 3) { //cntrl+c for copying data
				d = curr->data;
			}
			else if (ascii == 22) {//cntrl+v to paste data
				curr->data = d;
				printincell();
			}
			else if (ascii == 24) { //cntrl+x for cut data
				d = curr->data;
				curr->data="";
				system("cls");
				printGrid();
				printData();
				printcellborder(r_curr, c_curr, 4);
				printButton();
			}
			else if (ascii == 9) { //TAB to perform function
				functions();
			}
			/*else {
				char ch = ascii;
				curr->data = ch;
				printincell();
			}*/
			else
			{
				if (isprint(k) || k == ' ') {
					if (curr->data.size() < 4) {
						curr->data += k;
					}
					printincell();
				}
			}
		}
	}

	void Copy() {
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && temp->data != "") {
					v.push_back(temp->data);
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
	}
	void Cut() {
		Copy();
		for (int r = rs; r <= re; r++) {
			for (int c = cs; c <= ce; c++) {
				cell* temp = head;
				for (int i = 0; i < r; i++) 
					temp = temp->down;
				for (int j = 0; j < c; j++) 
					temp = temp->right;
				temp->data = "";
			}
		}
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}
	void Paste() {
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		int i = 0;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (i < v.size()) {
					temp->data = v[i++];
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		v.clear();
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}
	int sum() {
		int sum = 0;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && temp->data != "") {
					sum += stoi(temp->data);
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return sum;
	}
	int average() {
		int sum = 0;
		int count = 0;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && temp->data != "") {
					sum += stoi(temp->data);
				}
				count++;
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return sum / count;
	}
	int count() {
		int count = 0;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && temp->data != "") {
					count++;
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return count;
	}
	int Min() {
		int min = INT_MAX;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && temp->data != "") {
					int val = stoi(temp->data);
					if (val < min) {
						min = val;
					}
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return min;
	}
	int Max() {
		int max = INT_MIN;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && temp->data != "") {
					int val = stoi(temp->data);
					if (val > max) {
						max = val;
					}
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return max;
	}


	void saveToFile(const string& filename) {
		ofstream wrt(filename);
		if (wrt.is_open()) {
			cell* temp = head;
			for (int r = 0; r < r_size; r++) {
				cell* temp2 = temp;
				for (int c = 0; c < c_size; c++) {
					wrt << temp->data;
					if (c < c_size - 1)
						wrt << " ";
					temp = temp->right;
				}
				wrt << endl;
				temp = temp2->down;
			}
			wrt.close();
		}
	}
	void loadFromFile(const string& filename) {
		ifstream file(filename);
		if (file.is_open()) {
			clearData();
			string line;
			int r = 0;
			while (getline(file, line)) {
				stringstream ss(line);
				string cellData;
				int c = 0;
				while (getline(ss, cellData, ' ')) {
					if (r < r_size && c < c_size) {
						cell* temp = head;
						for (int i = 0; i < r; i++)
							temp = temp->down;
						for (int j = 0; j < c; j++)
							temp = temp->right;
						temp->data = cellData;
					}
					c++;
				}
				r++;
			}
			file.close();
			printGrid();
			printData();
			printButton();
			printcellborder(r_curr, c_curr, 4);
		}
	}

	void clearData() {
		color(7);
		cell* temp = head;
		cell* storCurr = curr;
		int storCurrRow = r_curr;
		int storCurrCol = c_curr;
		r_curr = 0;
		while (temp != nullptr) {
			c_curr = 0;
			curr = temp;
			while (curr != nullptr) {
				curr->data = "";
				curr = curr->right;
				c_curr++;
			}
			temp = temp->down;
			r_curr++;
		}
		r_curr = storCurrRow;
		c_curr = storCurrCol;
		curr = storCurr;
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}

	void functions() {
		int rpos, cpos;
		getRowColbyLeftClick(rpos, cpos);
		rpos /= cellHeight;
		cpos /= cellWidth;
		if (rpos == 0 && cpos == 20) {
			printcellborder(0, 20, 4);
			selectrange();
			inputkeyboard();
			curr->data = to_string(sum());
			printincell();
			printcellborder(0, 20, 7);
		}
		if (rpos == 2 && cpos == 20) {
			printcellborder(2, 20, 4);
			selectrange();
			inputkeyboard();
			curr->data = to_string(average());
			printincell();
			printcellborder(2, 20, 7);
		}
		if (rpos == 4 && cpos == 20) {
			printcellborder(4, 20, 4);
			selectrange();
			inputkeyboard();
			curr->data = to_string(count());
			printincell();
			printcellborder(4, 20, 7);
		}
		if (rpos == 6 && cpos == 20) {
			printcellborder(6, 20, 4);
			selectrange();
			inputkeyboard();
			curr->data = to_string(Min());
			printincell();
			printcellborder(6, 20, 7);
		}
		if (rpos == 8 && cpos == 20) {
			printcellborder(8, 20, 4);
			selectrange();
			inputkeyboard();
			curr->data = to_string(Max());
			printincell();
			printcellborder(8, 20, 7);
		}
		if (rpos == 12 && cpos == 20) {
			printcellborder(12, 20, 4);
			delRow();
			printcellborder(12, 20, 7);
		}
		if (rpos == 0 && cpos == 22) {
			printcellborder(0, 22, 4);
			clearData();
			printcellborder(0, 22, 7);
		}
		if (rpos == 2 && cpos == 22) {
			printcellborder(2, 22, 4);
			saveToFile("save-load.txt");
			printcellborder(2, 22, 7);
		}
		if (rpos == 4 && cpos == 22) {
			printcellborder(4, 22, 4);
			loadFromFile("save-load.txt");
			printcellborder(4, 22, 7);
		}
		if (rpos == 6 && cpos == 22) {
			printcellborder(6, 22, 4);
			selectrange();
			Cut();
			printcellborder(6, 22, 7);
		}
		if (rpos == 8 && cpos == 22) {
			printcellborder(8, 22, 4);
			selectrange();
			Copy();
			printcellborder(8, 22, 7);
		}
		if (rpos == 10 && cpos == 22) {
			printcellborder(10, 22, 4);
			selectrange();
			Paste();
			printcellborder(10, 22, 7);
		}
		if (rpos == 12 && cpos == 22) {
			printcellborder(12, 22, 4);
			delCol();
			printcellborder(12, 22, 7);
		}
	}
	void selectrange() {
		int ascii = 0; int selected = 0;
		while (selected < 2) {
			char k = _getch();
			ascii = k;
			if (ascii == -32) {
				char c = _getch();
				ascii = c;

				if (ascii == 77) {  //right arrow
					if (curr->right != nullptr) {
						printcellborder(r_curr, c_curr, 7);
						curr = curr->right;
						c_curr++;
						printcellborder(r_curr, c_curr, 4);
					}
				}
				else if (ascii == 75) {  //left arrow
					printcellborder(r_curr, c_curr, 7);
					if (curr->left != nullptr) {
						curr = curr->left;
						c_curr--;
					}
					printcellborder(r_curr, c_curr, 4);
				}
				else if (ascii == 72) {  //up arrow
					printcellborder(r_curr, c_curr, 7);
					if (curr->up != nullptr) {
						curr = curr->up;
						r_curr--;
					}
					printcellborder(r_curr, c_curr, 4);

				}
				else if (ascii == 80) {  //down arrow
					printcellborder(r_curr, c_curr, 7);
					if (curr->down != nullptr) {
						curr = curr->down;
						r_curr++;
					}
					printcellborder(r_curr, c_curr, 4);

				}
			}
			else if (ascii == 32) {  //space bar for selection of cell in range
				if (selected == 0) {
					rangestart = curr;
					rs = r_curr;
					cs = c_curr;
					printincell(18, 20);
				}
				if (selected == 1) {
					rangeend = curr;
					re = r_curr;
					ce = c_curr;
					printincell(18, 22);
				}
				selected++;
			}
			else {
				char ch = ascii;
				curr->data = ch;
				printincell();
			}
		}
	}
	void printButton() {
		int r = 1, c = 20;
//////////////////////
		printcellborder(0, 20);
		r = r * cellHeight -1;
		c = c * cellWidth +3;
		/*printcellborder(1, 10);
		r = r * cellHeight + 2;
		c = c * cellWidth + 5;*/
		gotoRowCol(r, c);
		cout << "Sum";

		printcellborder(2, 20);
		r = r +2* cellHeight;
		gotoRowCol(r, c);
		cout << "Avg";

		printcellborder(4, 20);
		r = r + 2 * cellHeight;
		gotoRowCol(r, c - 1);
		cout << "Count";

		printcellborder(6, 20);
		r = r + 2 * cellHeight;
		gotoRowCol(r, c);
		cout << "Min";

		printcellborder(8, 20);
		r = r + 2 * cellHeight;
		gotoRowCol(r, c);
		cout << "Max";

		printcellborder(12, 20);
		r = r + 2 * cellHeight;
		gotoRowCol(r+4, c-1);
		cout << "D Row";
/////////////////////

		r = 1;
		printcellborder(0, 22);
		r = r * cellHeight - 1;
		c = c + 2* cellWidth;
		gotoRowCol(r, c-1);
		cout << "Clear";

		printcellborder(2, 22);
		r = r + 2 * cellHeight;
		gotoRowCol(r, c-1);
		cout << "Save";

		printcellborder(4, 22);
		r = r + 2 * cellHeight;
		gotoRowCol(r, c - 1);
		cout << "Load";

		printcellborder(6, 22);
		r = r + 2 * cellHeight;
		gotoRowCol(r, c);
		cout << "Cut";


		printcellborder(8, 22);
		r = r + 2 * cellHeight;
		gotoRowCol(r, c-1);
		cout << "Copy";


		printcellborder(10, 22);
		r = r + 2 * cellHeight;
		gotoRowCol(r, c-2);
		cout << "Paste";

		printcellborder(12, 22);
		r = r + 2 * cellHeight;
		gotoRowCol(r, c - 1);
		cout << "D Col";

		r = 18, c = 20;
		//////////////////////
		r = r * cellHeight - 1;
		c = c * cellWidth-1;
		gotoRowCol(r, c);
		cout << "Range Start:"<<endl;
		
		c = 22;
		c = c * cellWidth -1;
		gotoRowCol(r, c);
		cout << "Range End:"<<endl;
		printcellborder(18, 20);
		printcellborder(18, 22);

	}
};


int main() {
	myExcel e;
	gotoRowCol(50, 0);

	/*while (1) {
		char c = _getch();
		int ascii = c;
		cout <<c<<" : " << ascii << endl;
	}*/

	//TAB to perform function
	//up/down/right/left arrows for movement
	//insert row col by pressing ENTER (w/a/s/d)
	//BACKSPACE for clearing rows/cols (r/c)
	//cntrl+c for copying data
	//cntrl+v to paste data
	//cntrl+x for cut data of one cell
	return 0;
}
