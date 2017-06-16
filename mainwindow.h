#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

class QAction;
class QLabel;
class FindDialog;
class Spreadsheet;

class MainWindow: public QMainWindow
{
	Q_OBJECT
public:
	MainWindow();
protected:
	void closeEvent(QCloseEvent * event);
private slots:
	void newFile();
	void open();
	bool save();
	bool saveAs();
	void find();
	void gotoCell();
	void sort();
	void about();

	void openRecentFile();
	void updateStatusBar();
	void spreadsheetModified();

private:
	void createActions();
	void createMenus();
	void createContextMenus();
	void createToolBars();
	void createStatusBar();
	void readSettings();
	void writeSettings();
	bool okToContinue();
	bool loadFile(const QString &filename);
	bool saveFile(const QString &filename);
	void setCurrentFile(const QString &filename);
	void updateRecentFileActions();
	QString strippedName(const QString &fullFileName);

	Spreadsheet *spreadsheet;
	FindDialog *findDialog;
	QLabel *locationLabel;
	QLabel *formulaLabel;
	QStringList recentFiles;
	QString curFile;

	enum { MaxRecentFiles = 5};
	QAction *recentFileActions[MaxRecentFiles];

	QMenu *fileMenu;
	QMenu *editMenu;
	QMenu *selectSubMenu;
	QMenu *toolsMenu;
	QMenu *optionsMenu;
	QMenu *helpMenu;

	QToolBar *fileToolBar;
	QToolBar *editToolBar;

	QAction *newAction;
	QAction *openAction;
	QAction *saveAction;
	QAction *saveAsAction;
	QAction *exitAction;
	QAction *selectAllAction;
	QAction *showGridAction;
	QAction *aboutAction;
	QAction *aboutQtAction;

	QAction *separatorAction;

	QAction *cutAction;
	QAction *copyAction;
	QAction *pasteAction;
	QAction *deleteAction;

	QAction *selectRowAction;
	QAction *selectColumnAction;
	QAction *findAction;
	QAction *goToCellAction;

	QAction *recalculateAction;
	QAction *sortAction;

	QAction *autoRecalcAction;
};


#endif