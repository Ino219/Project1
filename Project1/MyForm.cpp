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
	////Excelブックの表示はしない
	//app_->Visible = false;
	////ファイルパスからブックを開く
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
	////一枚目のワークシートを開く
	//worksheet = (Microsoft::Office::Interop::Excel::Worksheet^)workbook->Worksheets[1];
	//Range^ range = worksheet->Range[worksheet->Cells[1,1], worksheet->Cells[3, 3]];
	//range->Copy(Type::Missing);

	//String^ path3 = "C:\\Users\\chach\\Desktop\\new3.pptx";
	////パワーポイントCOM
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
	
	
	//段数
	int scaleNum = 2;
	//1段あたりの行数
	int rowMax = 17;
	//データ数
	int rowMax2 = rowMax*scaleNum;
	//列数
	int columnMax = 3;
	//テンプレート配置ページ
	int slideIndex = 1;
	//データ数
	int dataNum;
	//段組み数
	int dan;
	//試行失敗フラグ
	bool exit=false;
	//表数の最大値
	int maxScale=4;

	//データ作成
	cli::array<String^, 2>^ text = gcnew cli::array<String^, 2>(rowMax2,columnMax);
	//データ数分の反復処理
	for (int i = 0; i < rowMax2; i++) {
		//列数分の反復処理
		for (int j = 0; j < columnMax; j++) {
			Random^ rnd = gcnew Random();
			int ans = rnd->Next(10);
			//乱数を用いてデータを作成
			if(ans==i||ans==j){
				text[i, j] = "test:" + i + ":" + j + "\r\ntest2\r\ntest3";
			}
			else {
				text[i, j] = "test:" + i + ":" + j;
			}
		}
	}

	//データ数を取得
	dataNum = text->GetLength(0);
	//全てを1行分のデータとして、これがおさまるものを選ぶ
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

	//テンプレパスの指定
	String^ path = "C:\\Users\\chach\\Desktop\\new2.pptx";
	String^ path2 = "C:\\Users\\chach\\Desktop\\new2_after.pptx";

	
	//データ元ファイルのオープン
	Microsoft::Office::Interop::PowerPoint::Presentation^ presense = presen->Open(
		path,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse
	);
	//保存先の新規ファイルの作成
	Microsoft::Office::Interop::PowerPoint::Presentation^ presense2 = presen->Add(MsoTriState::msoFalse);

	//int maxrow = 0;
	//最初の表の番号
	int shapeNum = 1;

	//最初に追加するのは1番の表
	presense->SlideMaster->CustomLayouts[4]->Shapes[shapeNum]->Copy();
	//追加先にスライドを追加
	presense2->Slides->Add(1, Microsoft::Office::Interop::PowerPoint::PpSlideLayout::ppLayoutBlank);
	//追加したスライドにコピーしたオブジェクト(表)を貼付
	presense2->Slides[1]->Shapes->Paste();
	//追加先のスライドの図形の総数を取得
	int count=presense2->Slides[1]->Shapes->Count;
	//直近でコピーペーストした表を取得
	Microsoft::Office::Interop::PowerPoint::Table^ tab = presense2->Slides[1]->Shapes[count]->Table;
	

	//貼り付け時の表サイズ
	int standardTable = presense2->Slides[1]->Shapes[count]->Height;
	//別表への移行時、データのオフセットをかける
	int add = 0;
	//表のスライドの高さを規定
	//int tableHeight = (int)presense->PageSetup->SlideHeight-100;

	int tableHeight = standardTable+ presense2->Slides[1]->Shapes[count]->Top;
	//行の高さ
	int cellHeight = 0;
	//算出する表の現在の高さ
	int totalHeight = presense2->Slides[1]->Shapes[count]->Top;
	//int ct = 1;
	//MessageBox::Show(tableHeight.ToString());
	bool secondFlg = false;

	System::IO::StreamWriter^ sw = gcnew System::IO::StreamWriter(".\\logfile_new.txt");

	//データ数分の列を反復処理
	for (int i = 0; i < rowMax; i++) {
		//表が二つ目以降ならiを0に初期化
		if (secondFlg) {
			i = 0;
			secondFlg = false;
		}
		//行反復ごとにセルの高さを初期化
		cellHeight = 0;
		for (int j = 0; j < columnMax; j++) {

			//値の代入
			tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->Text = text[i+add, j];
			//フォントサイズ
			tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->Font->Size = 11;
			//フォントの色の指定
			tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->Font->Color->RGB = Color::Black.ToArgb();
			//上マージン
			tab->Cell(i + 1, j + 1)->Shape->TextFrame->MarginTop = 5;
			//下マージン
			tab->Cell(i + 1, j + 1)->Shape->TextFrame->MarginBottom = 5;



			//if (cellHeight < tab->Cell(i + 1, j + 1)->Shape->Height) {
			//行反復ごとにセルの高さを更新
			if (cellHeight < tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->BoundHeight) {

				//セルの図形の高さ(こちらは何故か更新されない)
				//cellHeight = tab->Cell(i + 1, j + 1)->Shape->Height;

				//MessageBox::Show(cellHeight + ":" + tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->BoundHeight);
				
				//テキストフレームの高さ(こちらは改行時、更新される)
				//10は余白分として足す
				cellHeight=tab->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->BoundHeight+10;
			}

			//最終行のイベント
			if ((i==rowMax-1)&&(j==columnMax-1)) {
				//図形のカウントを増やす
				shapeNum++;
				//次の図形のカウントが最大値を超えていたら、処理を中断
				if (shapeNum > maxScale) {
					exit = true;
					MessageBox::Show("break");
					break;
				}
				//次の表を追加
				presense->SlideMaster->CustomLayouts[4]->Shapes[shapeNum]->Copy();
				//追加したスライドにコピーしたオブジェクト(表)を貼付
				presense2->Slides[1]->Shapes->Paste();
				//追加先のスライドの図形の総数を取得
				int count = presense2->Slides[1]->Shapes->Count;

				//直近でコピーペーストした表を取得
				tab = presense2->Slides[1]->Shapes[count]->Table;
				add += rowMax;
				i = 0;
				j = 0;
				continue;
			}

		}
		
		//積算の高さを計算
			totalHeight += cellHeight;
			sw->Write("total:" + totalHeight + "cell:" + cellHeight);
			sw->WriteLine();
			sw->Write("rows:"+ tab->Rows[i+1]->Height);
			sw->WriteLine();
			//オーバーすれば、次の表へと移行
			if (totalHeight > tableHeight) {
				//MessageBox::Show(presense2->Slides[1]->Shapes[count]->Height.ToString());
				MessageBox::Show("bOer:"+totalHeight+"::"+tableHeight);
				//図形のカウントを増やす
				shapeNum++;
				//次の図形のカウントが最大値を超えていたら、処理を中断
				if (shapeNum > maxScale) {
					exit = true;
					MessageBox::Show("break:"+i);
					break;
				}
				//次の表をコピー
				presense->SlideMaster->CustomLayouts[4]->Shapes[shapeNum]->Copy();
				//追加したスライドにコピーしたオブジェクト(表)を貼付
				presense2->Slides[1]->Shapes->Paste();
				//追加先のスライドの図形の総数を取得
				int count = presense2->Slides[1]->Shapes->Count;
				//直近でコピーペーストした表を取得
				tab = presense2->Slides[1]->Shapes[count]->Table;
				//データの整合性を取るためのオフセット
				add += i;
				//ループ添字の初期化
				//i = 0;
				//表の高さの初期化
				totalHeight = presense2->Slides[1]->Shapes[count]->Top;
				//二つ目以降の表追加フラグの設定
				secondFlg = true;
				continue;
			}
	}

	sw->Close();

	//2番目に追加するのは4番
	//presense->SlideMaster->CustomLayouts[4]->Shapes[4]->Copy();
	//presense2->Slides[1]->Shapes->Paste();


	//テーブルを選択
	//Microsoft::Office::Interop::PowerPoint::Shape^ s = presense->Slides[2]->Shapes[1];
	//Microsoft::Office::Interop::PowerPoint::Shapes^ s2 = presense->Slides[2]->Shapes;
	//Microsoft::Office::Interop::PowerPoint::Shape^ s3;

	//表の名前を指定して、取得
	//for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in s2)
	//{
	//	if (var->Name == "表1") {
	//		//MessageBox::Show("表の取得");
	//		s3 = var;
	//	}
	//}

	//Microsoft::Office::Interop::PowerPoint::Table^ table = s->Table;

	//MessageBox::Show(table->Rows->Count.ToString());
	////図形に表が含まれていれば、メッセージを表示
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

		////スライドの中を探す
		//for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in s2)
		//{
		//	//表1から始める
		//	if (var->Name == "表"+ct) {
		//		//次があれば、表の差し替え
		//		s3 = var;
		//		nextTable = true;
		//	}
		//}

		//if (!nextTable) {
		//	//処理は中断
		//	exit = true;
		//	//break;
		//}

		//int cellHeight = 0;
		//int totalHeight = 0;
		//
		//for (int i = index; i < rowMax; i++) {
		//	for (int j = 0; j < columnMax; j++) {
		//		text[i, j] = "test:" + i + ":" + j + "\r\n" + "test2";
		//		//値の代入
		//		s3->Table->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->Text = text[i, j];
		//		
		//		if (cellHeight < s3->Table->Cell(i + 1, j + 1)->Shape->Height) {
		//			//セルの図形の高さ(こちらは何故か更新されない)
		//			cellHeight = s3->Table->Cell(i + 1, j + 1)->Shape->Height;
		//			//テキストフレームの高さ(こちらは改行時、更新される)
		//			cellHeight=s3->Table->Cell(i + 1, j + 1)->Shape->TextFrame->TextRange->BoundHeight;
		//		}
		//		
		//	}
		//	//積算の高さを計算
		//	totalHeight += cellHeight;
		//	//MessageBox::Show(tableHeight.ToString());
		//	//MessageBox::Show(totalHeight.ToString());
		//	//オーバーすれば、次の表へと移行
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

	//閉じる
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

	Excelのプロセスを閉じる処理
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
