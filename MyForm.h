#include <windows.h>
#include <msclr/marshal_cppstd.h>
#include <iostream>
#include <string>
using namespace System;
using namespace System::Windows::Forms;
using namespace System::Data;
using namespace System::Drawing;
using namespace System::Data::OleDb;

public ref class MainForm : public Form {
public:
    MainForm() {
        this->Text = "Список клієнтів страхової компанії";
        this->Size = System::Drawing::Size(800, 600);

        MenuStrip^ menuStrip = gcnew MenuStrip();
        ToolStripMenuItem^ aboutMenuItem = gcnew ToolStripMenuItem("Про програму");
        ToolStripMenuItem^ exitMenuItem = gcnew ToolStripMenuItem("Вихід");
        ToolStripMenuItem^ databaseMenuItem = gcnew ToolStripMenuItem("База даних");
        ToolStripMenuItem^ loadMenuItem = gcnew ToolStripMenuItem("Завантажити");
        ToolStripMenuItem^ addMenuItem = gcnew ToolStripMenuItem("Додати");
        ToolStripMenuItem^ updateMenuItem = gcnew ToolStripMenuItem("Оновити");
        ToolStripMenuItem^ deleteMenuItem = gcnew ToolStripMenuItem("Видалити");

        databaseMenuItem->DropDownItems->Add(loadMenuItem);
        databaseMenuItem->DropDownItems->Add(addMenuItem);
        databaseMenuItem->DropDownItems->Add(updateMenuItem);
        databaseMenuItem->DropDownItems->Add(deleteMenuItem);

        menuStrip->Items->Add(databaseMenuItem);
        menuStrip->Items->Add(aboutMenuItem);
        menuStrip->Items->Add(exitMenuItem);

        aboutMenuItem->Click += gcnew EventHandler(this, &MainForm::About_Click);
        exitMenuItem->Click += gcnew EventHandler(this, &MainForm::Exit_Click);
        loadMenuItem->Click += gcnew EventHandler(this, &MainForm::buttonLoad_Click);
        addMenuItem->Click += gcnew EventHandler(this, &MainForm::buttonAdd_Click);
        updateMenuItem->Click += gcnew EventHandler(this, &MainForm::buttonEd_Click);
        deleteMenuItem->Click += gcnew EventHandler(this, &MainForm::buttonDel_Click);

        this->MainMenuStrip = menuStrip;
        this->Controls->Add(menuStrip);

        Button^ btnLoad = gcnew Button();
        btnLoad->Text = "Завантажити";
        btnLoad->Location = Point(50, 50);
        btnLoad->Click += gcnew EventHandler(this, &MainForm::buttonLoad_Click);

        Button^ btnAdd = gcnew Button();
        btnAdd->Text = "Додати";
        btnAdd->Location = Point(50, 100);
        btnAdd->Click += gcnew EventHandler(this, &MainForm::buttonAdd_Click);

        Button^ btnUpdate = gcnew Button();
        btnUpdate->Text = "Оновити";
        btnUpdate->Location = Point(50, 150);
        btnUpdate->Click += gcnew EventHandler(this, &MainForm::buttonEd_Click);

        Button^ btnDelete = gcnew Button();
        btnDelete->Text = "Видалити";
        btnDelete->Location = Point(50, 200);
        btnDelete->Click += gcnew EventHandler(this, &MainForm::buttonDel_Click);

        this->Controls->Add(btnLoad);
        this->Controls->Add(btnAdd);
        this->Controls->Add(btnUpdate);
        this->Controls->Add(btnDelete);

        dataGridView1 = gcnew DataGridView();
        dataGridView1->Location = Point(200, 50);
        dataGridView1->Size = System::Drawing::Size(550, 400);
        dataGridView1->Columns->Add("ClientID", "ClientID");
        dataGridView1->Columns->Add("FullName", "FullName");
        dataGridView1->Columns->Add("Email", "Email");
        dataGridView1->Columns->Add("PhoneNumber", "PhoneNumber");
        dataGridView1->Columns->Add("PolicyNumber", "PolicyNumber");
        dataGridView1->Columns->Add("InsuranceType", "InsuranceType");
        dataGridView1->Columns->Add("StartDate", "StartDate");
        dataGridView1->Columns->Add("EndDate", "EndDate");
        this->Controls->Add(dataGridView1);
    }

private:
    DataGridView^ dataGridView1;

    void About_Click(Object^ sender, EventArgs^ e) {
        MessageBox::Show("Список клієнтів страхової компанії\nАвтор: Кулініч Владислав", "Про програму", MessageBoxButtons::OK, MessageBoxIcon::Information);
    }

    void Exit_Click(Object^ sender, EventArgs^ e) {
        Application::Exit();
    }

    void buttonLoad_Click(System::Object^ sender, System::EventArgs^ e) {
        String^ connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

        try {
            dbConnection->Open();
            String^ query = "SELECT * FROM Clients";
            OleDbCommand^ dbCommand = gcnew OleDbCommand(query, dbConnection);
            OleDbDataReader^ dbReader = dbCommand->ExecuteReader();

            dataGridView1->Rows->Clear();

            if (dbReader->HasRows == false) {
                MessageBox::Show("Дані відсутні", "Помилка");
            }
            else {
                while (dbReader->Read()) {
                    dataGridView1->Rows->Add(dbReader["ClientID"], dbReader["FullName"], dbReader["Email"], dbReader["PhoneNumber"], dbReader["PolicyNumber"], dbReader["InsuranceType"], dbReader["StartDate"], dbReader["EndDate"]);
                }
            }

            dbReader->Close();
            dbConnection->Close();
        }
        catch (Exception^ ex) {
            MessageBox::Show("Помилка з'єднання з базою даних: " + ex->Message, "Помилка");
        }
    }

    void buttonAdd_Click(System::Object^ sender, System::EventArgs^ e) {
        if (dataGridView1->SelectedRows->Count != 1) {
            MessageBox::Show("Потрібно обрати один рядок", "Зверніть увагу");
            return;
        }
        int index = dataGridView1->SelectedRows[0]->Index;
        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[2]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[3]->Value == nullptr) {
            MessageBox::Show("Не всі дані є", "Зверніть увагу");
            return;
        }

        String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        String^ name = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
        String^ email = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
        String^ phone = dataGridView1->Rows[index]->Cells[3]->Value->ToString();
        String^ policy = dataGridView1->Rows[index]->Cells[4]->Value->ToString();
        String^ type = dataGridView1->Rows[index]->Cells[5]->Value->ToString();
        String^ start = dataGridView1->Rows[index]->Cells[6]->Value->ToString();
        String^ end = dataGridView1->Rows[index]->Cells[7]->Value->ToString();

        String^ connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
        dbConnection->Open();
        String^ query = "INSERT INTO Clients VALUES (" + id + ",'" + name + "','" + email + "','" + phone + "','" + policy + "','" + type + "','" + start + "','" + end + "')";
        OleDbCommand^ dbCommand = gcnew OleDbCommand(query, dbConnection);

        if (dbCommand->ExecuteNonQuery() != 1)
            MessageBox::Show("Помилка у виконанні", "Помилка");
        else
            MessageBox::Show("Дані додані", "Ок");

        dbConnection->Close();
    }

    void buttonEd_Click(System::Object^ sender, System::EventArgs^ e) {
        if (dataGridView1->SelectedRows->Count != 1) {
            MessageBox::Show("Потрібно обрати один рядок", "Зверніть увагу");
            return;
        }
        int index = dataGridView1->SelectedRows[0]->Index;
        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[2]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[3]->Value == nullptr) {
            MessageBox::Show("Не всі дані є", "Зверніть увагу");
            return;
        }

        String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        String^ name = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
        String^ email = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
        String^ phone = dataGridView1->Rows[index]->Cells[3]->Value->ToString();
        String^ policy = dataGridView1->Rows[index]->Cells[4]->Value->ToString();
        String^ type = dataGridView1->Rows[index]->Cells[5]->Value->ToString();
        String^ start = dataGridView1->Rows[index]->Cells[6]->Value->ToString();
        String^ end = dataGridView1->Rows[index]->Cells[7]->Value->ToString();

        String^ connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
        dbConnection->Open();
        String^ query = "UPDATE Clients SET FullName='" + name + "', Email='" + email + "', PhoneNumber='" + phone + "', PolicyNumber='" + policy + "', InsuranceType='" + type + "', StartDate='" + start + "', EndDate='" + end + "' WHERE ClientID=" + id;
        OleDbCommand^ dbCommand = gcnew OleDbCommand(query, dbConnection);

        if (dbCommand->ExecuteNonQuery() != 1)
            MessageBox::Show("Помилка у виконанні", "Помилка");
        else
            MessageBox::Show("Дані оновлені", "Ок");

        dbConnection->Close();
    }

    void buttonDel_Click(System::Object^ sender, System::EventArgs^ e) {
        if (dataGridView1->SelectedRows->Count != 1) {
            MessageBox::Show("Потрібно обрати один рядок", "Зверніть увагу");
            return;
        }
        int index = dataGridView1->SelectedRows[0]->Index;
        String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();

        String^ connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
        dbConnection->Open();
        String^ query = "DELETE FROM Clients WHERE ClientID=" + id;
        OleDbCommand^ dbCommand = gcnew OleDbCommand(query, dbConnection);

        if (dbCommand->ExecuteNonQuery() != 1)
            MessageBox::Show("Помилка у виконанні", "Помилка");
        else
            MessageBox::Show("Дані видалені", "Ок");

        dbConnection->Close();
    }
};

[STAThread]
int main(array<System::String^>^ args) {
    Application::EnableVisualStyles();
    Application::SetCompatibleTextRenderingDefault(false);
    Application::Run(gcnew MainForm());
    return 0;
}
