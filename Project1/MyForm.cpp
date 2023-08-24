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
	//テンプレート配置ページ
	int slideIndex = 1;
	//データ数
	int dataNum;
	//段組み数
	int dan;
	//試行失敗フラグ
	bool exit=false;

	//データ作成
	cli::array<String^, 2>^ text = gcnew cli::array<String^, 2>(rowMax,columnMax);
	
	for (int i = 0; i < rowMax; i++) {
		for (int j = 0; j < columnMax; j++) {
				text[i, j] = "test:"+i+":"+j;			
		}
	}

	//データ数を取得
	dataNum = text->GetLength(0);
	//全てを1行分のデータとして、これがおさまるものを選ぶ
	if (dataNum >= 17) {
		slideIndex = 1;
		dan = 1;
	}if (dataNum <= 18 && dataNum >= 34) {
		slideIndex = 2;
		dan = 2;
	}

	//パスの指定
	String^ path = "C:\\Users\\chach\\Desktop\\new2.pptx";
	//パワーポイントCOM
	Microsoft::Office::Interop::PowerPoint::Application^ app = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
	Microsoft::Office::Interop::PowerPoint::Presentations^ presen = app->Presentations;
	Microsoft::Office::Interop::PowerPoint::Presentation^ presense = presen->Open(
		path,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse
	);
	//テーブルを選択
	Microsoft::Office::Interop::PowerPoint::Shape^ s = presense->Slides[2]->Shapes[1];
	Microsoft::Office::Interop::PowerPoint::Shapes^ s2 = presense->Slides[2]->Shapes;
	Microsoft::Office::Interop::PowerPoint::Shape^ s3;

	//表の名前を指定して、取得
	for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in s2)
	{
		if (var->Name == "表1") {
			//MessageBox::Show("表の取得");
			s3 = var;
		}
	}

	Microsoft::Office::Interop::PowerPoint::Table^ table = s->Table;

	//MessageBox::Show(table->Rows->Count.ToString());
	////図形に表が含まれていれば、メッセージを表示
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

		//スライドの中を探す
		for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in s2)
		{
			//表1から始める
			if (var->Name == "表"+ct) {
				//次があれば、表の差し替え
				s3 = var;
				nextTable = true;
			}
		}

		if (!nextTable) {
			//処理は中断
			exit = true;
			break;
		}

		int cellHeight = 0;
		int totalHeight = 0;
		
		for (int i = index; i < rowMax; i++) {
			for (int j = 0; j < columnMax; j++) {
				text[i, j] = "test:" + i + ":" + j + "\r\n" + "test2";
				//値の代入
				s3->Table->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->Text = text[i, j];
				
				if (cellHeight < s3->Table->Cell(i + 1, j + 1)->Shape->Height) {
					//セルの図形の高さ(こちらは何故か更新されない)
					cellHeight = s3->Table->Cell(i + 1, j + 1)->Shape->Height;
					//テキストフレームの高さ(こちらは改行時、更新される)
					cellHeight=s3->Table->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->BoundHeight;
				}
				
			}
			//積算の高さを計算
			totalHeight += cellHeight;
			MessageBox::Show(tableHeight.ToString());
			MessageBox::Show(totalHeight.ToString());
			//オーバーすれば、次の表へと移行
			if (totalHeight > tableHeight) {
				MessageBox::Show("bOer");
				index = i;
				ct++;
				break;
			}
		}
	}



	presense->Save();

	//閉じる
	presense->Close();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presense);
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presen);

	app->Quit();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(app);
	return System::Void();
}
