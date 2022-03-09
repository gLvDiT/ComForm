#include "MyForm.h"
#using <System.dll>
#include <iostream>
#include <fstream>
#include <windows.h>
#include <conio.h>
#include "Threads.h"

using namespace System;
using namespace System::Windows::Forms;
using namespace System::IO::Ports;

[STAThread]
void Main() {
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);
	ComForm::MyForm MyForm;
	Application::Run(% MyForm);
}