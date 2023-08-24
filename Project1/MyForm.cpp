#include "MyForm.h"

using namespace Project1;
using namespace Microsoft::Office::Core;
using namespace Microsoft::Office::Interop::PowerPoint;


[STAThreadAttribute]

int main() {
	System::Windows::Forms::Application::Run(gcnew MyForm());
	return 0;
}

System::Void Project1::MyForm::MyForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	int rowMax = 17;
	int columnMax = 3;
	//�e���v���[�g�z�u�y�[�W
	int slideIndex = 1;
	//�f�[�^��
	int dataNum;
	//�i�g�ݐ�
	int dan;
	//���s���s�t���O
	bool exit=false;

	//�f�[�^�쐬
	cli::array<String^, 2>^ text = gcnew cli::array<String^, 2>(rowMax,columnMax);
	
	for (int i = 0; i < rowMax; i++) {
		for (int j = 0; j < columnMax; j++) {
				text[i, j] = "test:"+i+":"+j;			
		}
	}

	//�f�[�^�����擾
	dataNum = text->GetLength(0);
	//�S�Ă�1�s���̃f�[�^�Ƃ��āA���ꂪ�����܂���̂�I��
	if (dataNum >= 17) {
		slideIndex = 1;
		dan = 1;
	}if (dataNum <= 18 && dataNum >= 34) {
		slideIndex = 2;
		dan = 2;
	}

	//�p�X�̎w��
	String^ path = "C:\\Users\\chach\\Desktop\\new2.pptx";
	//�p���[�|�C���gCOM
	Microsoft::Office::Interop::PowerPoint::Application^ app = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
	Microsoft::Office::Interop::PowerPoint::Presentations^ presen = app->Presentations;
	Microsoft::Office::Interop::PowerPoint::Presentation^ presense = presen->Open(
		path,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse
	);
	//�e�[�u����I��
	Microsoft::Office::Interop::PowerPoint::Shape^ s = presense->Slides[2]->Shapes[1];
	Microsoft::Office::Interop::PowerPoint::Shapes^ s2 = presense->Slides[2]->Shapes;
	Microsoft::Office::Interop::PowerPoint::Shape^ s3;

	//�\�̖��O���w�肵�āA�擾
	for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in s2)
	{
		if (var->Name == "�\1") {
			//MessageBox::Show("�\�̎擾");
			s3 = var;
		}
	}

	Microsoft::Office::Interop::PowerPoint::Table^ table = s->Table;

	//MessageBox::Show(table->Rows->Count.ToString());
	////�}�`�ɕ\���܂܂�Ă���΁A���b�Z�[�W��\��
	//if (s->HasTable==MsoTriState::msoTrue) {
	//	MessageBox::Show(s->Name);
	//}
	//int cellHeight = 0;
	//int totalHeight = 0;
	int tableHeight = (int)presense->PageSetup->SlideHeight;

	int index = 0;

	//MessageBox::Show(s3->Table->Cell(1, 1)->Shape->Height.ToString());
	//MessageBox::Show(s3->Table->Cell(2, 2)->Shape->Height.ToString());

	bool fFlag = false;
	int ct = 1;
	while (!fFlag) {

		bool nextTable = false;

		//�X���C�h�̒���T��
		for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in s2)
		{
			//�\1����n�߂�
			if (var->Name == "�\"+ct) {
				//��������΁A�\�̍����ւ�
				s3 = var;
				nextTable = true;
			}
		}

		if (!nextTable) {
			//�����͒��f
			exit = true;
			break;
		}

		int cellHeight = 0;
		int totalHeight = 0;
		
		for (int i = index; i < rowMax; i++) {
			for (int j = 0; j < columnMax; j++) {
				text[i, j] = "test:" + i + ":" + j + "\r\n" + "test2";
				//�l�̑��
				s3->Table->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->Text = text[i, j];
				
				if (cellHeight < s3->Table->Cell(i + 1, j + 1)->Shape->Height) {
					//�Z���̐}�`�̍���(������͉��̂��X�V����Ȃ�)
					cellHeight = s3->Table->Cell(i + 1, j + 1)->Shape->Height;
					//�e�L�X�g�t���[���̍���(������͉��s���A�X�V�����)
					cellHeight=s3->Table->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->BoundHeight;
				}
				
			}
			//�ώZ�̍������v�Z
			totalHeight += cellHeight;
			MessageBox::Show(tableHeight.ToString());
			MessageBox::Show(totalHeight.ToString());
			//�I�[�o�[����΁A���̕\�ւƈڍs
			if (totalHeight > tableHeight) {
				MessageBox::Show("bOer");
				index = i;
				ct++;
				break;
			}
		}
	}



	presense->Save();

	//����
	presense->Close();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presense);
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presen);

	app->Quit();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(app);
	return System::Void();
}
