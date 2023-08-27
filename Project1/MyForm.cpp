#include "MyForm.h"

using namespace Project1;
using namespace Microsoft::Office::Core;
using namespace Microsoft::Office::Interop::PowerPoint;
using namespace Microsoft::Office::Interop::Excel;

[STAThreadAttribute]

int main() {
	System::Windows::Forms::Application::Run(gcnew MyForm());
	return 0;
}

System::Void Project1::MyForm::MyForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	//String^ xlpath = "C:\\Users\\chach\\Desktop\\edc1.xlsx";

	//Microsoft::Office::Interop::Excel::Application^ app_ = nullptr;
	//Microsoft::Office::Interop::Excel::Workbook^ workbook = nullptr;
	//Microsoft::Office::Interop::Excel::Worksheet^ worksheet = nullptr;
	//Microsoft::Office::Interop::Excel::Range^ testRange = nullptr;
	//Microsoft::Office::Interop::Excel::ListObject^ lo = nullptr;

	//app_ = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();
	////Excel�u�b�N�̕\���͂��Ȃ�
	//app_->Visible = false;
	////�t�@�C���p�X����u�b�N���J��
	//workbook = (Microsoft::Office::Interop::Excel::Workbook^)(app_->Workbooks->Open(
	//	xlpath,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing,
	//	Type::Missing));
	////�ꖇ�ڂ̃��[�N�V�[�g���J��
	//worksheet = (Microsoft::Office::Interop::Excel::Worksheet^)workbook->Worksheets[1];
	//Range^ range = worksheet->Range[worksheet->Cells[1,1], worksheet->Cells[3, 3]];
	//range->Copy(Type::Missing);

	//String^ path3 = "C:\\Users\\chach\\Desktop\\new3.pptx";
	////�p���[�|�C���gCOM
	Microsoft::Office::Interop::PowerPoint::Application^ app = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
	Microsoft::Office::Interop::PowerPoint::Presentations^ presen = app->Presentations;
	//Microsoft::Office::Interop::PowerPoint::Presentation^ presense3 = presen->Open(
	//	path3,
	//	MsoTriState::msoFalse,
	//	MsoTriState::msoFalse,
	//	MsoTriState::msoTrue
	//);
	//Microsoft::Office::Interop::PowerPoint::Table^ tab2 = presense3->Slides[1]->Shapes[1]->Table;
	//tab2->Cell(1, 1)->Shape->Select(MsoTriState::msoCTrue);
	//presense3->Slides[1]->Shapes->Paste();
	
	
	//�i��
	int scaleNum = 2;
	//1�i������̍s��
	int rowMax = 17;
	//�f�[�^��
	int rowMax2 = rowMax*scaleNum;
	//��
	int columnMax = 3;
	//�e���v���[�g�z�u�y�[�W
	int slideIndex = 1;
	//�f�[�^��
	int dataNum;
	//�i�g�ݐ�
	int dan;
	//���s���s�t���O
	bool exit=false;
	//�\���̍ő�l
	int maxScale=4;

	//�f�[�^�쐬
	cli::array<String^, 2>^ text = gcnew cli::array<String^, 2>(rowMax2,columnMax);
	//�f�[�^�����̔�������
	for (int i = 0; i < rowMax2; i++) {
		//�񐔕��̔�������
		for (int j = 0; j < columnMax; j++) {
			Random^ rnd = gcnew Random();
			int ans = rnd->Next(10);
			//������p���ăf�[�^���쐬
			if(ans==i||ans==j){
				text[i, j] = "test:" + i + ":" + j + "\r\ntest2\r\ntest3";
			}
			else {
				text[i, j] = "test:" + i + ":" + j;
			}
		}
	}

	//�f�[�^�����擾
	dataNum = text->GetLength(0);
	//�S�Ă�1�s���̃f�[�^�Ƃ��āA���ꂪ�����܂���̂�I��
	if (dataNum >= 17) {
		slideIndex = 1;
		dan = 1;
	}
	else if (dataNum <= 18 && dataNum >= 34) {
		slideIndex = 2;
		dan = 2;
	}
	else if (dataNum <= 18 && dataNum >= 34) {
		slideIndex = 3;
		dan = 3;
	}
	else if (dataNum <= 18 && dataNum >= 34) {
		slideIndex = 4;
		dan = 4;
	}

	//�e���v���p�X�̎w��
	String^ path = "C:\\Users\\chach\\Desktop\\new2.pptx";
	String^ path2 = "C:\\Users\\chach\\Desktop\\new2_after.pptx";

	
	//�f�[�^���t�@�C���̃I�[�v��
	Microsoft::Office::Interop::PowerPoint::Presentation^ presense = presen->Open(
		path,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse
	);
	//�ۑ���̐V�K�t�@�C���̍쐬
	Microsoft::Office::Interop::PowerPoint::Presentation^ presense2 = presen->Add(MsoTriState::msoFalse);

	//int maxrow = 0;
	//�ŏ��̕\�̔ԍ�
	int shapeNum = 1;

	//�ŏ��ɒǉ�����̂�1�Ԃ̕\
	presense->SlideMaster->CustomLayouts[4]->Shapes[shapeNum]->Copy();
	//�ǉ���ɃX���C�h��ǉ�
	presense2->Slides->Add(1, Microsoft::Office::Interop::PowerPoint::PpSlideLayout::ppLayoutBlank);
	//�ǉ������X���C�h�ɃR�s�[�����I�u�W�F�N�g(�\)��\�t
	presense2->Slides[1]->Shapes->Paste();
	//�ǉ���̃X���C�h�̐}�`�̑������擾
	int count=presense2->Slides[1]->Shapes->Count;
	//���߂ŃR�s�[�y�[�X�g�����\���擾
	Microsoft::Office::Interop::PowerPoint::Table^ tab = presense2->Slides[1]->Shapes[count]->Table;
	

	//�\��t�����̕\�T�C�Y
	int standardTable = presense2->Slides[1]->Shapes[count]->Height;
	//�ʕ\�ւ̈ڍs���A�f�[�^�̃I�t�Z�b�g��������
	int add = 0;
	//�\�̃X���C�h�̍������K��
	//int tableHeight = (int)presense->PageSetup->SlideHeight-100;

	int tableHeight = standardTable+ presense2->Slides[1]->Shapes[count]->Top;
	//�s�̍���
	int cellHeight = 0;
	//�Z�o����\�̌��݂̍���
	int totalHeight = presense2->Slides[1]->Shapes[count]->Top;
	//int ct = 1;
	//MessageBox::Show(tableHeight.ToString());
	bool secondFlg = false;

	System::IO::StreamWriter^ sw = gcnew System::IO::StreamWriter(".\\logfile_new.txt");

	//�f�[�^�����̗�𔽕�����
	for (int i = 0; i < rowMax; i++) {
		//�\����ڈȍ~�Ȃ�i��0�ɏ�����
		if (secondFlg) {
			i = 0;
			secondFlg = false;
		}
		//�s�������ƂɃZ���̍�����������
		cellHeight = 0;
		for (int j = 0; j < columnMax; j++) {

			//�l�̑��
			tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->Text = text[i+add, j];
			//�t�H���g�T�C�Y
			tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->Font->Size = 11;
			//�t�H���g�̐F�̎w��
			tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->Font->Color->RGB = Color::Black.ToArgb();
			//��}�[�W��
			tab->Cell(i + 1, j + 1)->Shape->TextFrame->MarginTop = 5;
			//���}�[�W��
			tab->Cell(i + 1, j + 1)->Shape->TextFrame->MarginBottom = 5;



			//if (cellHeight < tab->Cell(i + 1, j + 1)->Shape->Height) {
			//�s�������ƂɃZ���̍������X�V
			if (cellHeight < tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->BoundHeight) {

				//�Z���̐}�`�̍���(������͉��̂��X�V����Ȃ�)
				//cellHeight = tab->Cell(i + 1, j + 1)->Shape->Height;

				//MessageBox::Show(cellHeight + ":" + tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->BoundHeight);
				
				//�e�L�X�g�t���[���̍���(������͉��s���A�X�V�����)
				//10�͗]�����Ƃ��đ���
				cellHeight=tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->BoundHeight+10;
			}

			//�ŏI�s�̃C�x���g
			if ((i==rowMax-1)&&(j==columnMax-1)) {
				//�}�`�̃J�E���g�𑝂₷
				shapeNum++;
				//���̐}�`�̃J�E���g���ő�l�𒴂��Ă�����A�����𒆒f
				if (shapeNum > maxScale) {
					exit = true;
					MessageBox::Show("break");
					break;
				}
				//���̕\��ǉ�
				presense->SlideMaster->CustomLayouts[4]->Shapes[shapeNum]->Copy();
				//�ǉ������X���C�h�ɃR�s�[�����I�u�W�F�N�g(�\)��\�t
				presense2->Slides[1]->Shapes->Paste();
				//�ǉ���̃X���C�h�̐}�`�̑������擾
				int count = presense2->Slides[1]->Shapes->Count;

				//���߂ŃR�s�[�y�[�X�g�����\���擾
				tab = presense2->Slides[1]->Shapes[count]->Table;
				add += rowMax;
				i = 0;
				j = 0;
				continue;
			}

		}
		
		//�ώZ�̍������v�Z
			totalHeight += cellHeight;
			sw->Write("total:" + totalHeight + "cell:" + cellHeight);
			sw->WriteLine();
			sw->Write("rows:"+ tab->Rows[i+1]->Height);
			sw->WriteLine();
			//�I�[�o�[����΁A���̕\�ւƈڍs
			if (totalHeight > tableHeight) {
				//MessageBox::Show(presense2->Slides[1]->Shapes[count]->Height.ToString());
				MessageBox::Show("bOer:"+totalHeight+"::"+tableHeight);
				//�}�`�̃J�E���g�𑝂₷
				shapeNum++;
				//���̐}�`�̃J�E���g���ő�l�𒴂��Ă�����A�����𒆒f
				if (shapeNum > maxScale) {
					exit = true;
					MessageBox::Show("break:"+i);
					break;
				}
				//���̕\���R�s�[
				presense->SlideMaster->CustomLayouts[4]->Shapes[shapeNum]->Copy();
				//�ǉ������X���C�h�ɃR�s�[�����I�u�W�F�N�g(�\)��\�t
				presense2->Slides[1]->Shapes->Paste();
				//�ǉ���̃X���C�h�̐}�`�̑������擾
				int count = presense2->Slides[1]->Shapes->Count;
				//���߂ŃR�s�[�y�[�X�g�����\���擾
				tab = presense2->Slides[1]->Shapes[count]->Table;
				//�f�[�^�̐���������邽�߂̃I�t�Z�b�g
				add += i;
				//���[�v�Y���̏�����
				//i = 0;
				//�\�̍����̏�����
				totalHeight = presense2->Slides[1]->Shapes[count]->Top;
				//��ڈȍ~�̕\�ǉ��t���O�̐ݒ�
				secondFlg = true;
				continue;
			}
	}

	sw->Close();

	//2�Ԗڂɒǉ�����̂�4��
	//presense->SlideMaster->CustomLayouts[4]->Shapes[4]->Copy();
	//presense2->Slides[1]->Shapes->Paste();


	//�e�[�u����I��
	//Microsoft::Office::Interop::PowerPoint::Shape^ s = presense->Slides[2]->Shapes[1];
	//Microsoft::Office::Interop::PowerPoint::Shapes^ s2 = presense->Slides[2]->Shapes;
	//Microsoft::Office::Interop::PowerPoint::Shape^ s3;

	//�\�̖��O���w�肵�āA�擾
	//for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in s2)
	//{
	//	if (var->Name == "�\1") {
	//		//MessageBox::Show("�\�̎擾");
	//		s3 = var;
	//	}
	//}

	//Microsoft::Office::Interop::PowerPoint::Table^ table = s->Table;

	//MessageBox::Show(table->Rows->Count.ToString());
	////�}�`�ɕ\���܂܂�Ă���΁A���b�Z�[�W��\��
	//if (s->HasTable==MsoTriState::msoTrue) {
	//	MessageBox::Show(s->Name);
	//}
	//int cellHeight = 0;
	//int totalHeight = 0;
	//int tableHeight = (int)presense->PageSetup->SlideHeight;

	//int index = 0;

	//MessageBox::Show(s3->Table->Cell(1, 1)->Shape->Height.ToString());
	//MessageBox::Show(s3->Table->Cell(2, 2)->Shape->Height.ToString());

	//bool fFlag = false;
	//int ct = 1;
	//while (!fFlag) {

		//bool nextTable = false;

		////�X���C�h�̒���T��
		//for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in s2)
		//{
		//	//�\1����n�߂�
		//	if (var->Name == "�\"+ct) {
		//		//��������΁A�\�̍����ւ�
		//		s3 = var;
		//		nextTable = true;
		//	}
		//}

		//if (!nextTable) {
		//	//�����͒��f
		//	exit = true;
		//	//break;
		//}

		//int cellHeight = 0;
		//int totalHeight = 0;
		//
		//for (int i = index; i < rowMax; i++) {
		//	for (int j = 0; j < columnMax; j++) {
		//		text[i, j] = "test:" + i + ":" + j + "\r\n" + "test2";
		//		//�l�̑��
		//		s3->Table->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->Text = text[i, j];
		//		
		//		if (cellHeight < s3->Table->Cell(i + 1, j + 1)->Shape->Height) {
		//			//�Z���̐}�`�̍���(������͉��̂��X�V����Ȃ�)
		//			cellHeight = s3->Table->Cell(i + 1, j + 1)->Shape->Height;
		//			//�e�L�X�g�t���[���̍���(������͉��s���A�X�V�����)
		//			cellHeight=s3->Table->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->BoundHeight;
		//		}
		//		
		//	}
		//	//�ώZ�̍������v�Z
		//	totalHeight += cellHeight;
		//	//MessageBox::Show(tableHeight.ToString());
		//	//MessageBox::Show(totalHeight.ToString());
		//	//�I�[�o�[����΁A���̕\�ւƈڍs
		//	if (totalHeight > tableHeight) {
		//		MessageBox::Show("bOer");
		//		index = i;
		//		ct++;
		//		break;
		//	}
		//}
	//}
	


	presense->Save();
	presense2->SaveAs(path2, Microsoft::Office::Interop::PowerPoint::PpSaveAsFileType::ppSaveAsDefault, MsoTriState::msoTrue);
	//presense3->Save();

	//����
	presense->Close();
	presense2->Close();
	//presense3->Close();

	//System::Runtime::InteropServices::Marshal::ReleaseComObject(presense3);

	System::Runtime::InteropServices::Marshal::ReleaseComObject(presense2);

	System::Runtime::InteropServices::Marshal::ReleaseComObject(presense);

	System::Runtime::InteropServices::Marshal::ReleaseComObject(presen);

	app->Quit();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(app);

	/*workbook->Save();

	Excel�̃v���Z�X����鏈��
	if (lo != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(lo);
		lo = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();
	if (testRange != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(testRange);
		testRange = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();
	workbook->Close(Type::Missing, Type::Missing, Type::Missing);
	if (workbook != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(workbook);
		workbook = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();
	if (app_ != nullptr)
	{
		app_->Quit();
		System::Runtime::InteropServices::Marshal::ReleaseComObject(app_);
		app_ = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();*/

	return System::Void();
}
