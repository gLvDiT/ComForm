#pragma once
#include <windows.h>
#include <ios>
#pragma hdrstop
#include <iostream> 
#include <fstream>
#include <sstream>
#include <cstring>
#include <stdio.h>
#include <conio.h> 
#include <msclr/marshal.h>
#include <atlstr.h>
#include <vcclr.h>
//Работа с класом баз данных CDataSource Access
#include <atldbcli.h>
#include <time.h>
//директория
#include <filesystem>

namespace ComForm {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Data;
	using namespace System::Windows::Forms;
	using namespace System::Windows::Forms;
	using namespace System::Drawing;
	using namespace System::IO::Ports;
	using namespace System::Threading;
	using namespace System::Text;
	using namespace System::Collections;
	using namespace std;
	using namespace System::Runtime::InteropServices;
	using namespace msclr::interop;
	using namespace msclr;
	using namespace System::Threading;
	using namespace System::Data::OleDb;

	public ref class MyForm : public System::Windows::Forms::Form
	{

		/////////////////////////////////////////////////////////////////////////////////////////////
	public:
		//Поток в глобальную переменную
		Thread^ SerialThRead; //Поток чтения сом порта
		Thread^ SerialThReadDataSP; //Поток расшифровки данных
		//Переменная кодировки Unicode
		//UnicodeEncoding^ unicode = gcnew UnicodeEncoding;
		//ASCIIEncoding^ ascii = gcnew ASCIIEncoding;
		//Масиив для байтов сом порта
		cli::array<Byte>^ array_SerialBytes;
		cli::array<Char>^ array_SerialChar;
		//Строка для раскодировки байтов
		String^ unicodeString;
		//Нумератор для цикла раскодировки
		IEnumerator^ myEnum;
		String^ text_index = "false";
		//Переменная для перобразования String^ в const char
		marshal_context^ context;// = gcn;
		cli::array<Char>^ str_bin_array = gcnew cli::array<Char>(44);

		//LPCTSTR source = L".\\db.mdb";
		//C:\ATS_ComForm
		LPCTSTR source = L"C:\\ATS_ComForm\\db.mdb";

		//Работа с БД
		OleDbConnection^ connect = gcnew OleDbConnection(L"data source=C:\\ATS_ComForm\\db.mdb;provider=microsoft.jet.oledb.4.0;");
		//Текстовый файл для записи выбранного СОМ порта

	private: System::Windows::Forms::NotifyIcon^ notifyIcon1;
	private: System::Windows::Forms::NotifyIcon^ notifyIcon2;
	private: System::Windows::Forms::TextBox^ textBoxHistory;

	private: System::Windows::Forms::Label^ label4;


	bool dirExists(const std::string& dirName_in)
	{
		DWORD ftyp = GetFileAttributesA(dirName_in.c_str());
		if (ftyp == INVALID_FILE_ATTRIBUTES)
			return false;  // ошибка

		if (ftyp & FILE_ATTRIBUTE_DIRECTORY)
			return true;   // существует

		return false;    // не является дир.
	}

	public:

		bool source_bool;

		MyForm(void)
		{
			InitializeComponent();
			//Отключаем контроль передачи параметров на форму из других потоков
			//метод не очень состоятелен, возможны тормоза или форма начинает лагать.
			//или в потоке необходима пауза минимум в 100 мл сек.
			Control::CheckForIllegalCrossThreadCalls = false;

			//Поищем текстовый файл с запомненным СОМ портом
			HANDLE txtFind = INVALID_HANDLE_VALUE;
			WIN32_FIND_DATA txtFileData;

			if (dirExists("C:\\ATS_ComForm") == false) 
			{
				std::filesystem::create_directories("C:\\ATS_ComForm");
			}

			txtFind = FindFirstFile(_T("C:\\ATS_ComForm\\dbCom.txt"), &txtFileData);
			//Если txt не нашли, то создаём новый
			if (txtFind == INVALID_HANDLE_VALUE)
			{
				FILE* fp = fopen("C:\\ATS_ComForm\\dbCom.txt", "w+");
				fclose(fp);
			};

			//Поищем файл БД Access db.mdb
			HANDLE hFind = INVALID_HANDLE_VALUE;
			WIN32_FIND_DATA pFileData;
			hFind = FindFirstFile(_T("C:\\ATS_ComForm\\db.mdb"), &pFileData);
			//Если БД не нашли, то создаём базу данных Acces
			if (hFind == INVALID_HANDLE_VALUE)
			{
				FindClose(hFind);
				//Создаём файл БД в корне программы
				source_bool = CreateAccessDatabase(source);
				if (source_bool = true)
				{
					//Создаём таблицу в БД
					int i; //Устанавливаем 132 поля c избытком (132 байта в бинарном виде)
					String^ Fill_bd = "CREATE TABLE TABLE1 (id COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, S_Date DATE, ";

					Fill_bd += "SumByte TEXT(22), ";
					Fill_bd += "Header22 TEXT(22), ";
					Fill_bd += "Header44 TEXT(22), ";
					Fill_bd += "T1_2 TEXT(22), ";
					Fill_bd += "T2_2 TEXT(22), ";
					Fill_bd += "T1_6 TEXT(22), ";
					Fill_bd += "T2_6 TEXT(22), ";
					Fill_bd += "Posled22 TEXT(22), ";
					Fill_bd += "Posled44 TEXT(22), ";
					Fill_bd += "ContSumm22 TEXT(22), ";
					Fill_bd += "ContSumm44 TEXT(22), ";
					Fill_bd += "ExtOUT TEXT(22), ";
					Fill_bd += "ExtIN TEXT(22), ";
					Fill_bd += "A22 TEXT(22), ";
					Fill_bd += "A44 TEXT(22), ";
					Fill_bd += "B22 TEXT(22), ";
					Fill_bd += "B44 TEXT(22), ";
					Fill_bd += "B222 TEXT(22), ";
					Fill_bd += "B444 TEXT(22), ";
					Fill_bd += "S_Hour TEXT(22), ";
					Fill_bd += "S_Minute TEXT(22), ";
					Fill_bd += "S_Second TEXT(22), ";
					Fill_bd += "NumberDay TEXT(22), ";
					Fill_bd += "DurationSeconds TEXT(22), ";
					Fill_bd += "Code44 TEXT(22), ";
					Fill_bd += "NumberPhone TEXT(22), ";


					String^ end_string = "" ", ";
					for (i = 0; i <= 131; i++)
					{
						if (i == 131)
						{
							end_string = "";
						};
						Fill_bd += "Fill" + i + " TEXT(16)" + end_string;
					}
					Fill_bd += ")";
					connect->Open();
					System::Data::OleDb::OleDbCommand^ command = gcnew System::Data::OleDb::OleDbCommand(Fill_bd, connect);

					try
					{
						command->ExecuteNonQuery();
					}
					catch (OleDbException^ pe)
					{

						this->richTextBox1->AppendText(pe + "     " + Fill_bd);

					}
					connect->Close();
				}
			}

			findPorts();
			hFind = FindFirstFile(_T("C:\\ATS_ComForm\\db.mdb"), &pFileData);
			//Если база есть то сразу начинаем считывание
			if (hFind != INVALID_HANDLE_VALUE)
			{
				OpenCom();
			}

		}

	protected:

		~MyForm()
		{
			//Закрыть порт и поток чтения порта			
			if (components)
			{
				delete components;
			}
		}

	private:

	private: System::IO::Ports::SerialPort^ serialPort;
	private: System::Windows::Forms::ComboBox^ comboBox1;
	private: System::Windows::Forms::ComboBox^ comboBox2;
	private: System::Windows::Forms::Label^ label1;
	private: System::Windows::Forms::Label^ label2;
	private: System::Windows::Forms::Button^ button1;
	private: System::Windows::Forms::Button^ button2;
	private: System::Windows::Forms::TextBox^ textBox1;
	private: System::Windows::Forms::TextBox^ textBox2;
	private: System::Windows::Forms::Button^ button3;
	private: System::Windows::Forms::ProgressBar^ progressBar1;
	private: System::Windows::Forms::RichTextBox^ richTextBox1;
	private: System::Windows::Forms::Label^ label3;
	private: System::Windows::Forms::TextBox^ textBox_Indicator;
	private: System::ComponentModel::IContainer^ components;


#pragma region Windows Form Designer generated code
		   void InitializeComponent(void)
		   {
			   this->components = (gcnew System::ComponentModel::Container());
			   System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(MyForm::typeid));
			   this->serialPort = (gcnew System::IO::Ports::SerialPort(this->components));
			   this->comboBox1 = (gcnew System::Windows::Forms::ComboBox());
			   this->comboBox2 = (gcnew System::Windows::Forms::ComboBox());
			   this->label1 = (gcnew System::Windows::Forms::Label());
			   this->label2 = (gcnew System::Windows::Forms::Label());
			   this->button1 = (gcnew System::Windows::Forms::Button());
			   this->button2 = (gcnew System::Windows::Forms::Button());
			   this->textBox1 = (gcnew System::Windows::Forms::TextBox());
			   this->textBox2 = (gcnew System::Windows::Forms::TextBox());
			   this->button3 = (gcnew System::Windows::Forms::Button());
			   this->progressBar1 = (gcnew System::Windows::Forms::ProgressBar());
			   this->richTextBox1 = (gcnew System::Windows::Forms::RichTextBox());
			   this->label3 = (gcnew System::Windows::Forms::Label());
			   this->textBox_Indicator = (gcnew System::Windows::Forms::TextBox());
			   this->notifyIcon1 = (gcnew System::Windows::Forms::NotifyIcon(this->components));
			   this->notifyIcon2 = (gcnew System::Windows::Forms::NotifyIcon(this->components));
			   this->textBoxHistory = (gcnew System::Windows::Forms::TextBox());
			   this->label4 = (gcnew System::Windows::Forms::Label());
			   this->SuspendLayout();
			   // 
			   // serialPort
			   // 
			   this->serialPort->ReadTimeout = 500;
			   this->serialPort->WriteTimeout = 500;
			   // 
			   // comboBox1
			   // 
			   this->comboBox1->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
			   this->comboBox1->FormattingEnabled = true;
			   this->comboBox1->Location = System::Drawing::Point(114, 21);
			   this->comboBox1->Name = L"comboBox1";
			   this->comboBox1->Size = System::Drawing::Size(159, 21);
			   this->comboBox1->TabIndex = 0;
			   this->comboBox1->SelectedIndexChanged += gcnew System::EventHandler(this, &MyForm::comboBox1_SelectedIndexChanged);
			   // 
			   // comboBox2
			   // 
			   this->comboBox2->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
			   this->comboBox2->FormattingEnabled = true;
			   this->comboBox2->Items->AddRange(gcnew cli::array< System::Object^  >(2) { L"9600", L"115200" });
			   this->comboBox2->Location = System::Drawing::Point(114, 48);
			   this->comboBox2->Name = L"comboBox2";
			   this->comboBox2->Size = System::Drawing::Size(159, 21);
			   this->comboBox2->TabIndex = 1;
			   // 
			   // label1
			   // 
			   this->label1->AutoSize = true;
			   this->label1->Location = System::Drawing::Point(3, 24);
			   this->label1->Name = L"label1";
			   this->label1->Size = System::Drawing::Size(62, 13);
			   this->label1->TabIndex = 2;
			   this->label1->Text = L"COM Порт:";
			   // 
			   // label2
			   // 
			   this->label2->AutoSize = true;
			   this->label2->Location = System::Drawing::Point(3, 51);
			   this->label2->Name = L"label2";
			   this->label2->Size = System::Drawing::Size(108, 13);
			   this->label2->TabIndex = 3;
			   this->label2->Text = L"Скорость передачи:";
			   this->label2->Click += gcnew System::EventHandler(this, &MyForm::label2_Click);
			   // 
			   // button1
			   // 
			   this->button1->Location = System::Drawing::Point(278, 21);
			   this->button1->Name = L"button1";
			   this->button1->Size = System::Drawing::Size(196, 23);
			   this->button1->TabIndex = 4;
			   this->button1->Text = L"Открыть порт и начать чтение";
			   this->button1->UseVisualStyleBackColor = true;
			   this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			   // 
			   // button2
			   // 
			   this->button2->Location = System::Drawing::Point(279, 48);
			   this->button2->Name = L"button2";
			   this->button2->Size = System::Drawing::Size(195, 23);
			   this->button2->TabIndex = 5;
			   this->button2->Text = L"Закрыть порт и прекратить чтение";
			   this->button2->UseVisualStyleBackColor = true;
			   this->button2->Click += gcnew System::EventHandler(this, &MyForm::button2_Click);
			   // 
			   // textBox1
			   // 
			   this->textBox1->Location = System::Drawing::Point(2, 130);
			   this->textBox1->Name = L"textBox1";
			   this->textBox1->ReadOnly = true;
			   this->textBox1->RightToLeft = System::Windows::Forms::RightToLeft::No;
			   this->textBox1->Size = System::Drawing::Size(765, 20);
			   this->textBox1->TabIndex = 6;
			   this->textBox1->Text = L"Сообщения";
			   // 
			   // textBox2
			   // 
			   this->textBox2->Location = System::Drawing::Point(5, 95);
			   this->textBox2->Name = L"textBox2";
			   this->textBox2->Size = System::Drawing::Size(619, 20);
			   this->textBox2->TabIndex = 7;
			   this->textBox2->Text = L"Ввести комманду";
			   // 
			   // button3
			   // 
			   this->button3->Location = System::Drawing::Point(630, 92);
			   this->button3->Name = L"button3";
			   this->button3->Size = System::Drawing::Size(126, 23);
			   this->button3->TabIndex = 8;
			   this->button3->Text = L"Отправить комманду";
			   this->button3->UseVisualStyleBackColor = true;
			   this->button3->Click += gcnew System::EventHandler(this, &MyForm::button3_Click);
			   // 
			   // progressBar1
			   // 
			   this->progressBar1->BackColor = System::Drawing::SystemColors::Control;
			   this->progressBar1->Location = System::Drawing::Point(5, 75);
			   this->progressBar1->Name = L"progressBar1";
			   this->progressBar1->Size = System::Drawing::Size(379, 11);
			   this->progressBar1->TabIndex = 11;
			   this->progressBar1->Click += gcnew System::EventHandler(this, &MyForm::progressBar1_Click);
			   // 
			   // richTextBox1
			   // 
			   this->richTextBox1->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
				   | System::Windows::Forms::AnchorStyles::Left)
				   | System::Windows::Forms::AnchorStyles::Right));
			   this->richTextBox1->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 14.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				   static_cast<System::Byte>(204)));
			   this->richTextBox1->Location = System::Drawing::Point(2, 156);
			   this->richTextBox1->Name = L"richTextBox1";
			   this->richTextBox1->Size = System::Drawing::Size(765, 361);
			   this->richTextBox1->TabIndex = 15;
			   this->richTextBox1->Text = L"";
			   this->richTextBox1->TextChanged += gcnew System::EventHandler(this, &MyForm::richTextBox1_TextChanged);
			   // 
			   // label3
			   // 
			   this->label3->AutoSize = true;
			   this->label3->BackColor = System::Drawing::Color::Red;
			   this->label3->Location = System::Drawing::Point(443, 24);
			   this->label3->Name = L"label3";
			   this->label3->Size = System::Drawing::Size(0, 13);
			   this->label3->TabIndex = 16;
			   // 
			   // textBox_Indicator
			   // 
			   this->textBox_Indicator->BackColor = System::Drawing::Color::Red;
			   this->textBox_Indicator->Location = System::Drawing::Point(71, 22);
			   this->textBox_Indicator->Name = L"textBox_Indicator";
			   this->textBox_Indicator->ReadOnly = true;
			   this->textBox_Indicator->Size = System::Drawing::Size(21, 20);
			   this->textBox_Indicator->TabIndex = 17;
			   // 
			   // notifyIcon1
			   // 
			   this->notifyIcon1->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"notifyIcon1.Icon")));
			   this->notifyIcon1->Text = L"Сборщик ComForm";
			   this->notifyIcon1->Visible = true;
			   this->notifyIcon1->MouseDoubleClick += gcnew System::Windows::Forms::MouseEventHandler(this, &MyForm::notifyIcon1_MouseDoubleClick);
			   // 
			   // notifyIcon2
			   // 
			   this->notifyIcon2->Text = L"notifyIcon2";
			   this->notifyIcon2->Visible = true;
			   // 
			   // textBoxHistory
			   // 
			   this->textBoxHistory->Location = System::Drawing::Point(489, 51);
			   this->textBoxHistory->Name = L"textBoxHistory";
			   this->textBoxHistory->ReadOnly = true;
			   this->textBoxHistory->RightToLeft = System::Windows::Forms::RightToLeft::No;
			   this->textBoxHistory->Size = System::Drawing::Size(267, 20);
			   this->textBoxHistory->TabIndex = 18;
			   this->textBoxHistory->Text = L"История";
			   this->textBoxHistory->TextChanged += gcnew System::EventHandler(this, &MyForm::textBox3_TextChanged);
			   // 
			   // label4
			   // 
			   this->label4->AutoSize = true;
			   this->label4->Location = System::Drawing::Point(486, 35);
			   this->label4->Name = L"label4";
			   this->label4->Size = System::Drawing::Size(149, 13);
			   this->label4->TabIndex = 19;
			   this->label4->Text = L"История последней записи:";
			   this->label4->Click += gcnew System::EventHandler(this, &MyForm::label4_Click);
			   // 
			   // MyForm
			   // 
			   this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			   this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			   this->AutoSize = true;
			   this->ClientSize = System::Drawing::Size(768, 518);
			   this->Controls->Add(this->label4);
			   this->Controls->Add(this->textBoxHistory);
			   this->Controls->Add(this->textBox_Indicator);
			   this->Controls->Add(this->label3);
			   this->Controls->Add(this->richTextBox1);
			   this->Controls->Add(this->progressBar1);
			   this->Controls->Add(this->button3);
			   this->Controls->Add(this->textBox2);
			   this->Controls->Add(this->textBox1);
			   this->Controls->Add(this->button2);
			   this->Controls->Add(this->button1);
			   this->Controls->Add(this->label2);
			   this->Controls->Add(this->label1);
			   this->Controls->Add(this->comboBox2);
			   this->Controls->Add(this->comboBox1);
			   this->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"$this.Icon")));
			   this->Name = L"MyForm";
			   this->Text = L"Сборщик АТС - ATS ComForm";
			   this->FormClosing += gcnew System::Windows::Forms::FormClosingEventHandler(this, &MyForm::FormCancel);
			   this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			   this->Resize += gcnew System::EventHandler(this, &MyForm::MyForm_Resize);
			   this->ResumeLayout(false);
			   this->PerformLayout();

		   }
#pragma endregion

		   //////////////////////////////////////////////////////////////////////////////////////
		   //Access создать новую базу
	private: BOOLEAN  CreateAccessDatabase(LPCTSTR szDatabasePath) //LPCTSTR szDatabasePath
	{
		CDataSource ds;
		IDBDataSourceAdmin* pIDBDataSourceAdmin = NULL;

		CLSID clsid = { 0xdee35070L, 0x506b, 0x11cf,
		{ 0xb1, 0xaa, 0x0, 0xaa, 0x0, 0xb8, 0xde, 0x95 } };
		HRESULT hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER,
			__uuidof(IDBInitialize), (void**)&ds.m_spInit);
		if (FAILED(hr))
			return false;

		USES_CONVERSION;

		CDBPropSet rgPropertySet(DBPROPSET_DBINIT);
		rgPropertySet.AddProperty(DBPROP_INIT_DATASOURCE,
			T2BSTR(szDatabasePath));

		hr = ds.m_spInit->QueryInterface(IID_IDBDataSourceAdmin,
			(void**)&pIDBDataSourceAdmin);
		if (FAILED(hr))
		{
			ds.Close();
			return false;
		}

		hr = pIDBDataSourceAdmin->CreateDataSource(1,
			&rgPropertySet, NULL, IID_NULL, NULL);
		if (FAILED(hr))
		{
			pIDBDataSourceAdmin->Release();
			ds.Close();
			return false;
		}

		pIDBDataSourceAdmin->Release();
		ds.Close();



		return true;
	}

		   //////////////////////////////////////////////////////////////////////////////////////
		   //Поиск портов, передача списка в массив и инициализация масива в комбобокс
	private: void findPorts(void)
	{
		// Найти порты по имени
		cli::array<Object^>^ objectArray = SerialPort::GetPortNames();
		// передать список в комбо

		this->comboBox1->Items->AddRange(objectArray);

		//Проверка наличия записи СОМ порта в тектовом файле
		char s[5]; //Значение порта из файла
		long file_size; //Размер полученных данных из файла если ноль то он пустой
		fstream f;
		f.open("C:\\ATS_ComForm\\dbCom.txt", ios_base::in);//открываем поток для чтения
		f.getline(s, 5); //Читаем файл в s
		f.seekg(0, ios::end); //Для правильного подсчета размера строки устанавливаем позицию чтения в 0
		file_size = f.tellg(); //Размер полученой строки
		f.close();
		//Преобразовываем char в System::String
		System::String^ strCLR = gcnew System::String(s);

		try
		{
			//this->comboBox1->Text =  this->comboBox1->Items[0]->ToString();

			this->comboBox1->Text = file_size == 0 ? this->comboBox1->Items[0]->ToString() : strCLR->ToString();

			this->comboBox2->Text = "9600";
		}
		catch (OleDbException^ pe)
		{

			this->richTextBox1->AppendText(pe + " - Не найден СОМ порт");

		}

	}

		   /////////////////////////////////////////////////////////////////////////////////////
		   //Открыть СОМ порт
	private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
		OpenCom();
	}

	private: void OpenCom() {

		this->textBox2->Text = String::Empty;
		if (this->comboBox1->Text == String::Empty || this->comboBox2->Text == String::Empty)
			this->textBox1->Text = "Выберите порт";
		else {
			try {
				//убедимся, что порт не открыт кем-то ещё	
				if (!this->serialPort->IsOpen) {
					this->serialPort->PortName = this->comboBox1->Text;
					this->serialPort->BaudRate = Int32::Parse(this->comboBox2->Text);
					this->textBox2->Text = "Текст сообщения";
					//Открыть порт 
					this->serialPort->Open();
					this->progressBar1->Value = 100;
					this->textBox_Indicator->BackColor = Color::DarkGreen;

					//Запускается поток чтения СОМ порта
					SerialThRead = gcnew Thread(gcnew ParameterizedThreadStart(this, &MyForm::ReadSerialPort));
					//Меняем приоритет фоновых потоков, т.е. при закрытии главного потока все фоновые потоки тоже закроются
					//(если это не сделать но форма закроется, а приложение продолжит работу)
					SerialThRead->IsBackground = true;
					SerialThRead->Start("SerialThRead1");
					this->textBox1->Text = "Идёт чтение с СОМ порта";
					this->richTextBox1->AppendText("Идёт чтение данных с СОМ порта" + System::Environment::NewLine);
					connect->Open();
					this->button1->Enabled = false;
					this->comboBox1->Enabled = false;
					this->comboBox2->Enabled = false;
					//Очистить файл
					FILE* fp = fopen("C:\\ATS_ComForm\\dbCom.txt", "w");
					fclose(fp);
					//Запись в файл  выбранного значения сом порта
					fstream f;
					f.open("C:\\ATS_ComForm\\dbCom.txt", ios_base::out);//открываем поток для записи
					f.write((char*)(void*)Marshal::StringToHGlobalAnsi(this->comboBox1->Text), 5); //Преобразовываем String в char b pfgbcsdftv d afqk
					f.close();

				}
				else
					this->textBox1->Text = "Порт не открывается";
			}
			catch (UnauthorizedAccessException^) {
				this->textBox1->Text = "Неразрешенный доступ";
			}
		}


	}

		   /////////////////////////////////////////////////////////////////////////////////////		
		   //Закрыть СОМ порт
	private: System::Void button2_Click(System::Object^ sender, System::EventArgs^ e) {
		String^ re = "";
		if (this->serialPort->IsOpen)
		{
			//Закрыть поток чтения СОМ порта
			SerialThRead->Abort();
			//Закрыть портt
			this->serialPort->Close();
			//Очистить файл с подлюкчением
			FILE* fp = fopen("C:\\ATS_ComForm\\dbCom.txt", "w");
			fclose(fp);

			connect->Close();
		}
		// обновить прогресс
		this->progressBar1->Value = 0;
		// Включить кнопку чтения
		this->button1->Enabled = true;
		this->comboBox1->Enabled = true;
		this->comboBox2->Enabled = true;
		// Включить кнопку открытия порта
		this->button3->Enabled = true;
		this->textBox_Indicator->BackColor = Color::Red;
		this->richTextBox1->AppendText("Чтение СОМ порта остановлено\n");
		this->textBox1->Text = "Чтение СОМ порта остановлено";
	}

		   /////////////////////////////////////////////////////////////////////////////////////
		   //Отправить сообшение
	private: System::Void button3_Click(System::Object^ sender, System::EventArgs^ e) {
		//Добавим имя отправителя
		String^ name = this->serialPort->PortName;
		//захватить текст и сохранить в буфер отправки
		String^ message = this->textBox2->Text;
		//передать в порт
		if (this->serialPort->IsOpen)
			this->serialPort->WriteLine(message);
		else
			this->textBox1->Text = "СОМ порт не открыт";
	}


		   /////////////////////////////////////////////////////////////////////////////////////
		   //Функция чтения СОМ порта (выполняется в потоке SerialThRead)
	private: System::Void ReadSerialPort(Object^ data)
	{
		int i = 0;
		if (this->serialPort->IsOpen)
		{
			while (true)//(this->serialPort->IsOpen)
			{
				try
				{
					//Определить массив и прочитать в него данные из порта

					if (this->serialPort->IsOpen) {
						if (serialPort->BytesToRead > 0)
						{
							// this->richTextBox1->AppendText("Байты получены");
							int Count_Byte = 0;
							SerialThRead->Sleep(500);
							array_SerialBytes = gcnew cli::array<Byte>(serialPort->BytesToRead);
							Count_Byte = serialPort->BytesToRead; //после чтения сом порта "BytesToRead" очиститься, поэтому читаем его заранее
							this->serialPort->Read(array_SerialBytes, 0, serialPort->BytesToRead);
							//SerialThRead->Sleep(100);
							ReadSerialPortDataSP(Count_Byte);
							//Запись полученных данных в БД

						}
					}
					//i.ToDecimal();
				}
				catch (TimeoutException^ ex)
				{
					//this->textBox1->Text = "Чтение СОМ порта";
					//MessageBox::Show("Ишибка чтения с СОМ порта", "Ишибка чтения с СОМ порта", MessageBoxButtons::OK, MessageBoxIcon::Asterisk);
					//this->textBox1->Text = ex->Message;
				}
			}
		}
	}

		   /////////////////////////////////////////////////////////////////////////////////////
		   //Функция преобразования полученных данных (выполняется во втором потоке SerialThReadDataSP)
	private: System::Void ReadSerialPortDataSP(int Count_Byte)
	{
		//while (true)
		//{
			//SerialThReadDataSP->Sleep(50);
		if (Count_Byte > 0)
		{

			String^ byte_String = "Строка байтов = ";
			String^ byte_String2 = "Строка байтов = ";//Контрольная переменная
			String^ binary_dec; //Переведённый dec в bin
			String^ bin_String = "";// Готовая строка двоичных кодов
			int dec_int;//Переведённый bin в dec для дальнейшего перевода в ACSII
			int i_num_int;//Нумератор последовательности полученых битов
			String^ Bin_Encode_String;// Готовая строка двоичных кодов
			String^ Select_String = "";// Строка подключения
			String^ Value_String = "";// Строка подключения
			//Формируем строку с полученными байтами
			IEnumerator^ myEnum = array_SerialBytes->GetEnumerator();
			//this->richTextBox1->AppendText(array_SerialBytes[0].ToString());
			String^ Date_String;// Дата для выявления времени ошибки записи
			//str_bin_array
			//bin_String = "";
			int i = 0;
			int Summ_Count_Byte = Count_Byte - 1;

			//connect->Open();
			//Формирование полей запроса инсерт
			Value_String = "VALUES(?,";
			Select_String = "INSERT INTO TABLE1 ([S_Date], ";
			//Формирование текста запроса
			for (i = 0; i <= 131; i++)
			{
				Select_String += "[Fill" + i + "], ";
				Value_String += "?,";
			}
			Select_String += "[SumByte], ";
			Select_String += "[ExtOUT], ";
			Select_String += "[ExtIN], ";
			Select_String += "[T1_2], ";
			Select_String += "[T2_2], ";
			Select_String += "[T1_6], ";
			Select_String += "[T2_6], ";
			Select_String += "[A22], ";
			Select_String += "[B22], ";
			Select_String += "[B222], ";
			Select_String += "[S_Hour], ";
			Select_String += "[S_Minute], ";
			Select_String += "[S_Second], ";
			Select_String += "[NumberDay], ";
			Select_String += "[DurationSeconds] ";
			Value_String += "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?";

			Select_String += ",[Header22], ";
			Select_String += "[Posled22], ";
			Select_String += "[ContSumm22] ";
			Value_String += ",?,?,?";

			Select_String += ",[Header44], ";
			Select_String += "[Posled44], ";
			Select_String += "[ContSumm44], ";
			Select_String += "[A44], ";
			Select_String += "[B44], ";
			Select_String += "[B444], ";
			Select_String += "[Code44], ";
			Select_String += "[NumberPhone] ";
			Value_String += ",?,?,?,?,?,?,?,?";

			Select_String += ")";
			Value_String += ")";
			Select_String += Value_String;

			OleDbCommand^ comm = gcnew OleDbCommand(Select_String, connect);

			comm->Parameters->Add("S_Date", OleDb::OleDbType::BSTR)->Value = DateTime::Today.Now.ToString();
			this->textBoxHistory->Text = DateTime::Today.Now.ToString();
			Date_String = DateTime::Today.Now.ToString();

			i = 0;
			//this->richTextBox1->AppendText("Начало цикла");
			while (myEnum->MoveNext())
			{
				Byte b = safe_cast<Byte>(myEnum->Current);
				//Десятичный вид
				byte_String = byte_String + "#" + (char)b;
				//int Int_dec = (b & 0x0FFF); 
				//Бинарный вид
				binary_dec = binEncodeByte(b);
				bin_String = bin_String + binary_dec;
				comm->Parameters->Add("Fill" + i, OleDb::OleDbType::BSTR)->Value = binary_dec;
				i++;
			}

			//this->richTextBox1->AppendText(byte_String);
			//Если полей меньше чем 132 то добавляем значения с "false"
			int num_b;
			num_b = i;
			for (i = num_b; i <= 131; i++)
			{
				comm->Parameters->Add("Fill" + i, OleDb::OleDbType::BSTR)->Value = text_index;
			}

			comm->Parameters->Add("SumByte", OleDb::OleDbType::BSTR)->Value = Count_Byte.ToString();
			comm->Parameters->Add("ExtOUT", OleDb::OleDbType::BSTR)->Value = (28 + 4) / 8 <= Count_Byte ?
				DecEncodeByte(bin_String->Substring(16, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(20, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(24, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(28, 4)).ToString() : text_index;

			comm->Parameters->Add("ExtIN", OleDb::OleDbType::BSTR)->Value = (60 + 4) / 8 <= Count_Byte ?
				DecEncodeByte(bin_String->Substring(48, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(52, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(56, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(60, 4)).ToString() : text_index;

			comm->Parameters->Add("T1_2", OleDb::OleDbType::BSTR)->Value = (16 + 1) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(16, 1)).ToString() : text_index;
			comm->Parameters->Add("T2_2", OleDb::OleDbType::BSTR)->Value = (17 + 1) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(17, 1)).ToString() : text_index;
			comm->Parameters->Add("T1_6", OleDb::OleDbType::BSTR)->Value = (48 + 1) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(48, 1)).ToString() : text_index;
			comm->Parameters->Add("T2_6", OleDb::OleDbType::BSTR)->Value = (49 + 1) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(49, 1)).ToString() : text_index;

			comm->Parameters->Add("A22", OleDb::OleDbType::BSTR)->Value = (40 + 2) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(40, 2)).ToString() : text_index;
			comm->Parameters->Add("B22", OleDb::OleDbType::BSTR)->Value = (46 + 2) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(46, 2)).ToString() : text_index;
			comm->Parameters->Add("B222", OleDb::OleDbType::BSTR)->Value = (72 + 2) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(72, 2)).ToString() : text_index;
			comm->Parameters->Add("S_Hour", OleDb::OleDbType::BSTR)->Value = (75 + 5) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(75, 5)).ToString() : text_index;
			comm->Parameters->Add("S_Minute", OleDb::OleDbType::BSTR)->Value = (82 + 6) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(82, 6)).ToString() : text_index;
			comm->Parameters->Add("S_Second", OleDb::OleDbType::BSTR)->Value = (90 + 6) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(90, 6)).ToString() : text_index;
			comm->Parameters->Add("NumberDay", OleDb::OleDbType::BSTR)->Value = (103 + 9) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(103, 9)).ToString() : text_index;

			comm->Parameters->Add("DurationSeconds", OleDb::OleDbType::BSTR)->Value = (112 + 16) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(112, 16)).ToString() : text_index;
			comm->Parameters->Add("Header22", OleDb::OleDbType::BSTR)->Value = (0 + 16) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(0, 16)).ToString() : text_index;
			comm->Parameters->Add("Posled22", OleDb::OleDbType::BSTR)->Value = (160 + 8) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(160, 8)).ToString() : text_index;
			comm->Parameters->Add("ContSumm22", OleDb::OleDbType::BSTR)->Value = (168 + 8) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(168, 8)).ToString() : text_index;

			comm->Parameters->Add("Header44", OleDb::OleDbType::BSTR)->Value = (176 + 16) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(176, 16)).ToString() : text_index;
			comm->Parameters->Add("Posled44", OleDb::OleDbType::BSTR)->Value = (336 + 8) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(336, 8)).ToString() : text_index;
			comm->Parameters->Add("ContSumm44", OleDb::OleDbType::BSTR)->Value = (344 + 8) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(344, 8)).ToString() : text_index;
			comm->Parameters->Add("A44", OleDb::OleDbType::BSTR)->Value = (222 + 2) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(222, 2)).ToString() : text_index;
			comm->Parameters->Add("B44", OleDb::OleDbType::BSTR)->Value = (216 + 2) / 8 <= Count_Byte ? DecEncodeByte(bin_String->Substring(216, 2)).ToString() : text_index;
			comm->Parameters->Add("B444", OleDb::OleDbType::BSTR)->Value = text_index;
			comm->Parameters->Add("Code44", OleDb::OleDbType::BSTR)->Value = (268 + 4) / 8 <= Count_Byte ?
				DecEncodeByte(bin_String->Substring(192, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(196, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(200, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(204, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(208, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(212, 4)).ToString() : text_index;

			comm->Parameters->Add("NumberPhone", OleDb::OleDbType::BSTR)->Value = (268 + 4) / 8 <= Count_Byte ?
				DecEncodeByte(bin_String->Substring(224, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(228, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(232, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(236, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(240, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(244, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(248, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(252, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(256, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(260, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(264, 4)).ToString() +
				DecEncodeByte(bin_String->Substring(268, 4)).ToString() : text_index;

			try
			{
				comm->ExecuteNonQuery();
				//this->richTextBox1->AppendText(L"Записано");

			}
			catch (OleDbException^ pe)
			{

				this->richTextBox1->AppendText(L"Ишибка записи в БД: " + Date_String + "  -----  " + bin_String + System::Environment::NewLine + pe);

			}
			delete byte_String, comm, binary_dec, myEnum, Select_String, bin_String, Date_String, Value_String, byte_String2, Summ_Count_Byte;

			Count_Byte = 0;
		}
		//}


	}

	private:  String^ binEncodebin(int i_num_int, String^ binary_dec)
	{
		//Формируем байты по правилам
		String^ binary_dec_res;
		binary_dec_res = binary_dec;
		return binary_dec_res;
	}

	private:  String^ binEncodeByte(char a)
	{
		String^ binary;
		binary = "";
		int i;
		for (i = 7; i >= 0; i--)
		{
			if ((a >> i) && (0x01 != 0))
				binary = binary + "1";
			else
				binary = binary + "0";
		}
		return binary;
	}

		   //Функция перевода двоичной в десячиную
		   //Вернуть sum
	private: System::Int32 DecEncodeByte(String^ str4)
	{
		//String^ S_Str4 = str4;


		//Преобразуем String в const char*
		if (str4 == "1010") { return  0; };
		marshal_context^ context = gcnew marshal_context();
		const char* binary = context->marshal_as<const char*>(str4);

		int len, dec = 0, i, exp;
		float two = 2;

		len = strlen(binary);
		exp = len - 1;

		for (i = 0; i < len; i++, exp--) {

			dec += binary[i] == '1' ? pow(two, exp) : 0;
		}

		puts(binary);
		delete context;
		return  dec;
	}


	private: System::Void comboBox1_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void label2_Click(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void MyForm_Load(System::Object^ sender, System::EventArgs^ e) {
	}

		   /* Закрытие программы через крестик */
	private: System::Void FormCancel(System::Object^ sender, FormClosingEventArgs^ e)
	{
		if (this->serialPort->IsOpen)
		{
			MessageBox::Show(L"Не закрыто соединение с СОМ портом!", L"Завершите соединение с СОМ портом");
			e->Cancel = true;
		}
	}

	private: System::Void progressBar1_Click(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void label3_Click(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void richTextBox1_TextChanged(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void ChB_synchron_CheckedChanged(System::Object^ sender, System::EventArgs^ e) {
	}

	private: System::Void button4_Click(System::Object^ sender, System::EventArgs^ e) {
	}

	private: System::Void notifyIcon1_MouseDoubleClick(System::Object^ sender, System::Windows::Forms::MouseEventArgs^ e) {
		this->Show();
		WindowState = FormWindowState::Normal;
		notifyIcon1->Visible = false;
	}
	private: System::Void MyForm_Resize(System::Object^ sender, System::EventArgs^ e) {
		if (this->WindowState == FormWindowState::Minimized)
		{
			this->Hide();
			notifyIcon1->Visible = true;
		}
	}
	private: System::Void label4_Click(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void textBox3_TextChanged(System::Object^ sender, System::EventArgs^ e) {
	}
	};
}


