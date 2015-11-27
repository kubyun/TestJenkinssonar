#include <FileBrowser/Docviewer/OfficePreComp.hpp>
//tes
//#include	"xlsevaluator.h"
//#include	"xlslocalinfo.h"
#include	"xlscalvalue.h"
//#include	"xlsstringparser.h"
//#include	"xlsbookiterator.h"
//#include	"xlsvaluecriteria.h"
//#include	"xlsvalueformat.h"
//#include	"xlsformatbuffer.h"
#include	"xlscalccell.h"
//#include	"xlscell.h"
//#include	"xlsrow.h"
//#include	"xlscharbuffer.h"
//#include	"xlssheet.h"
//#include	"xlsformula.h"
#include	"xlscalcalcengine.h"
//#include	"xlsdatabase.h"
//#include	"xlslistargsfunc.h"
//#include	"xlsnumberlistargsfunc.h"
//#include	"xlspwnumberlistargsfunc.h"
//#include	"xlsbook.h"
//#include	"xlssheet.h"
//#include	"xlsformula.h"
#include	"xlsToken.h"
//#include	"xlsgroup.h"
//#include	"xlsmath.h"
//#include	"xlscaldatabase.h"
#include	"xlsbondfuncs.h"
#include	"xlsengineerfuncs.h"
#include	"xlsmiscaddinfuncs.h"
#include	"xlstokensum.h"
#include	"xlstokenfunc.h"
#include	"xlsfunc.h"
#include	"xlsevaluator.h"
#include	"xlsArrayEvaluator.h"

#ifdef USE_ARRAYFUNCTION_COUTSOURCING

xlsArrayEvaluator::xlsArrayEvaluator(xlsEvaluator* evaluator) : BrBase()
{
	m_evaluator = evaluator;
	m_calcEngine = m_evaluator->m_calcEngine;

	// [��ļ���ó��-2] ��ļ��Ŀ����� ����� �����ϱ� ���� ���ÿ������ �ʱ�ȭ
	m_arrayResultVals = BrNULL;

	// [��ļ���ó��-2] ��ļ��Ŀ��꿡 ����Ǵ� �Ķ���͵��� ���� ���ú������� �ʱ�ȭ
	m_arrayInputVal1 = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
	m_arrayInputVal2 = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
	m_arrayInputVal3 = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
	m_arrayInputVal4 = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
	m_arrayInputVal5 = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);

	// main cell�� �ʱ�ȭ
	m_pCalcCell = BrNULL;
}

xlsArrayEvaluator::~xlsArrayEvaluator()
{
	// [��ļ���ó��-2] ��ļ��Ŀ��꿡 ����� �������� �ع�
	BR_SAFE_DELETE(m_arrayResultVals);
	BrDELETE m_arrayInputVal1;
	BrDELETE m_arrayInputVal2;
	BrDELETE m_arrayInputVal3;
	BrDELETE m_arrayInputVal4;
	BrDELETE m_arrayInputVal5;

	// ��ļ��Ŀ��꿡 ����� �������� �ع�
	if (m_arrayInputVals.GetSize() > 0) {
		for (int i = 0; i < m_arrayInputVals.GetSize(); i++) {
			xlsCalValue* pValue = m_arrayInputVals[i];
			BR_SAFE_DELETE(pValue);
		}
	}
}

void xlsArrayEvaluator::recalcArrayFormula(xlsCalcCell* cell)
{
	// cell����
	m_pCalcCell = cell;

	// ��ļ�������ɿ��� �Ǵ��ϱ� �� ��ļ����� ����cell�� ����Ǵ��� ����cell�� ����Ǵ��� �Ǵ��ϱ�.
	// ����cell�� ����� ��� �Է¹�İ� ��������� ũ�Ⱑ ���ƾ� �Ѵ�.
	BrBOOL bSignleCell = BrTRUE; // �������δ� ����cell�� ����Ǵ°����� ����.
	int nCols = cell->m_arrayRef.getNrCols();
	int nRows = cell->m_arrayRef.getNrRows();
	if (nCols > 1 || nRows > 1)
		bSignleCell = BrFALSE;
	else
		bSignleCell = BrTRUE;

	// [��ļ���ó��-2] 4Ģ������, �Լ� � ���ؼ��� �Է��Ķ���Ϳ� ��ļ�������ɿ��� ������ 
	// ó���� ������İ��� �ٸ��� ��ȯó���ǿ��� �Ѵ�.
	BR_SAFE_DELETE(m_arrayResultVals);
	m_arrayResultVals = BrNEW xlsValueArray();
	m_arrayResultVals->setSize(1, 1);

	xlsCalValue *val1 = NULL, *val2 = NULL;
	xlsToken *token = cell->m_formula->m_firstToken, *token_backup = NULL;
	while (token != NULL) {
		int nClass = token->getClass();
		switch(nClass) {
		case m_eXLS_TokenAdd:
		case m_eXLS_TokenDiv:
		case m_eXLS_TokenMul:
		case m_eXLS_TokenSub:
			{
				token = processNumericalExpression(token);
				continue;
			}
			// 3�ܰ�(�⺻�Լ�����)���� Sum��ɱ����� �����Ͽ� �Ʒ��� �ڵ�� ��ġ�Ƿ�
			// �Ʒ��� �ڵ带 �ּ�ó���Ѵ�.
			//case m_eXLS_TokenSum:
			//	{
			//		xlsTokenSum* pToken = (xlsTokenSum*)token;
			//		token = token->evaluate(m_evaluator);
			//		continue;
			//	}
		case m_eXLS_TokenFuncRand:
		case m_eXLS_TokenFuncBasic:
			{
				token = processTokenFuncVarBasic(token);
				continue;
			}
		case m_eXLS_TokenFuncSqrt:
		case m_eXLS_TokenFuncSign:
		case m_eXLS_TokenFuncExp:
		case m_eXLS_TokenFuncAbs:
		case m_eXLS_TokenFuncNormSDist:
		case m_eXLS_TokenFunc:
			{
				// token = processTokenFunc(token, nCount, bSignleCell);
				token = processTokenFunc(token);
				continue;
			}
			// �Լ������� ����Լ��� �����Ҽ� ���� �ǿ�����. ���� �� �Լ��� ���� �籸���� �ʿ���.
		case m_eXLS_TokenChoose:
			{
				token = doTokenChooseFuncVar(token);
				continue;
			}
		case m_eXLS_TokenSum:
		case m_eXLS_TokenFuncVarBasic:
		case m_eXLS_TokenFuncVar:
			{
				token = processTokenFuncVar(token);
				continue;
			}
		case m_eXLS_TokenIf:
			// 4�ܰ迡�� ó���ǿ��� �� ����.
			//token = doTokenFuncVar(token);
			//break;
		case m_eXLS_TokenGE:
		case m_eXLS_TokenEQ:
		case m_eXLS_TokenGT:
		case m_eXLS_TokenLE:
		case m_eXLS_TokenLT:
		case m_eXLS_TokenNE:
			// 4�ܰ迡�� ó���ǿ��� �� ����.
			token = token->m_next;
			break;
		default:
			token = token->evaluate(m_evaluator);
			continue;
		}
	}

	BR_SAFE_DELETE(m_arrayResultVals);
}

// [��ļ���ó��-2] ��ļ��Ŀ� ���� 4Ģ����ó��
// 4Ģ���꿡 ���� ���ó���̹Ƿ� �Է��Ķ������ ������ 2����°��� ������ �Ѵ�.
// token : 4Ģ���꿡 ���� token
// ������ : 4Ģ������ ���� token
xlsToken* xlsArrayEvaluator::processNumericalExpression(xlsToken* token)
{
	xlsCalValue *val1 = NULL, *val2 = NULL;
	xlsToken* token_backup = NULL;

	// �ʱ�ȭ
	val1 = m_evaluator->m_val->m_prev;
	val2 = m_evaluator->m_val;

	// ù��° �Ķ������ Backup
	m_arrayInputVal1->copy(val1);

	// �ι�° �Ķ������ Backup
	m_arrayInputVal2->copy(val2);

	// ���������� �����ϴ°��� ��������
	bool bCheck = checkArrayFromulaCondition(val1, val2);
	if (!bCheck) {
		(*m_evaluator->m_vals)[0]->setError(eNA);
		//		token = token->evaluate(this);
		return BrNULL;
	}

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nIndex < nTotalCount) {
		getValFromInputParameter(val1, m_arrayInputVal1, nIndex);
		getValFromInputParameter(val2, m_arrayInputVal2, nIndex);

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �缳��
		m_evaluator->m_val = val2;
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� �����(val1)�� �����ϱ�
	// �ֳ��ϸ� ��Ŀ����� ���� �Ķ���Ͱ� n���� ���(2���̻�)�� �ֱ⶧���̴�.
	int nVal = val1->m_nVal;
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// [��ļ���ó��-2] ��ļ����� ���������� �����Ǵ°��� �˻��ϴ� �Լ�
// val1 : ù��° �Է��Ķ����
// val2 : �ι�° �Է��Ķ����
// ������ : true-���Ǽ���, false-���Ǻ���
bool xlsArrayEvaluator::checkArrayFromulaCondition(xlsCalValue * val1, xlsCalValue * val2)
{
	// �ΰ��� �Ķ���Ͱ� ��� Ȥ�� �ɿ��ΰ��� �˻��ϱ�
	bool bFlag1 = (val1->isRange() || val1->isArray());
	bool bFlag2 = (val2->isRange() || val2->isArray());

	// ù��° �Ķ������ ������ ���
	int nRows1 = 0, nCols1 = 0;
	if (val1->isRange()) {
		xlsTRange rng;
		val1->getRange(rng);
		nRows1 = rng.getNrRows();
		nCols1 = rng.getNrCols();
	}
	else if (val1->isArray()) {
		xlsValueArray* va = val1->m_array;
		nRows1 = va->getRowCount();
		nCols1 = va->getColCount();
	}
	else {
		nRows1 = 0;
		nCols1 = 0;
	}

	// �ι�° �Ķ������ ������ ���
	int nRows2 = 0, nCols2 = 0;
	if (val2->isRange()) {
		xlsTRange rng;
		val2->getRange(rng);
		nRows2 = rng.getNrRows();
		nCols2 = rng.getNrCols();
	}
	else if (val2->isArray()) {
		xlsValueArray* va = val2->m_array;
		nRows2 = va->getRowCount();
		nCols2 = va->getColCount();
	}
	else {
		nRows2 = 0;
		nCols2 = 0;
	}

	//�ΰ��� �Ķ������ ��� �ϳ��� ��� Ȥ�� �ɿ��� �ƴ϶��
	if (!bFlag1 || !bFlag2) {
		// �ΰ��� �Ķ������ �ϳ��� ����̶�� �� ����� ũ�⸦ �˾Ƴ���
		int nRows = 0, nCols = 0;
		if (bFlag1) {
			nRows = nRows1;
			nCols = nCols1;
		}
		else if (bFlag2) {
			nRows = nRows2;
			nCols = nCols2;
		}
		else {
			// Empty Process
		}

		if (nRows == 0 && nCols == 0)
			m_arrayResultVals->setSize(1, 1);
		else
			m_arrayResultVals->setSize(nRows, nCols);

		return true;
	}

	// �� �Ķ���Ϳ� ���� ��
	if (nRows1 == nRows2 && nCols1 == nCols2) {
		if (nRows1 > m_arrayResultVals->getRowCount() || nCols1 > m_arrayResultVals->getColCount())
			m_arrayResultVals->setSize(nRows1, nCols1);

		return true;
	}
	else {
		// 2���� �μ��� �� 1��������̶��
		if (nRows1 == 1 && nRows1 == nCols2) {
			m_arrayResultVals->setSize(nRows2, nCols1);
			return true;
		}

		if (nCols1 == 1 && nCols1 == nRows2) {
			m_arrayResultVals->setSize(nRows1, nCols2);
			return true;
		}

		// row������ ���� col������ 1�� ������迡 �ִٸ�
		int nMaxCol = BrMAX(nCols1, nCols2);
		int nMinCol = BrMIN(nCols1, nCols2);
		if (nRows1 == nRows2 && nMinCol == 1 && (nMaxCol % nMinCol) == 0) {
			m_arrayResultVals->setSize(nRows1, nMaxCol);
			return true;
		}

		// col������ ���� row������ 1�� ������迡 �ִٸ�
		int nMaxRow = BrMAX(nRows1, nRows2);
		int nMinRow = BrMIN(nRows1, nRows2);
		if (nCols1 == nCols2 && nMinRow == 1 && (nMaxRow % nMinRow) == 0) {
			m_arrayResultVals->setSize(nMaxRow, nCols1);
			return true;
		}

		return false;
	}
}

// [��ļ���ó��-2] ��ļ�������ɿ��� ũ�⸦ result������� ũ��� �����ϱ� 
void xlsArrayEvaluator::setResultBuffer()
{
	int nRows = 0, nCols = 0;
	xlsCalcCell* pCell = m_pCalcCell;
	xlsTRange rng;
	nRows = pCell->m_arrayRef.getNrRows();
	nCols = pCell->m_arrayRef.getNrCols();

	m_arrayResultVals->setSize(nRows, nCols);
}

// [��ļ���ó��-2] �⺻�Լ��� ����Լ����� �־��� �Է��Ķ���Ϳ��� nIndex�� �ش��� ���� ���
// pDst : ����� ���� �����ϴ� ����
// pSrc : �Է��Ķ����
// nIndex : �Է��Ķ������ ����
// ������ : �ش� token�� ���� token
void xlsArrayEvaluator::getValInFunc(xlsCalValue* pDst, xlsCalValue* pSrc, int nIndex)
{
	// ����Ŀ� ������ �� �� �Ĺ�ȣ�� ���
	int c_rows = m_arrayResultVals->getRowCount(); // cell�׷��� �ళ��(�Ƿ�, A1:D10���� v_rows = 10)
	int c_cols = m_arrayResultVals->getColCount(); // cell�׷��� �İ���(�Ƿ�, A1:D10���� v_cols = 4)
	int a_r = nIndex / c_cols;
	int a_c = nIndex % c_cols;

	if (pSrc->isRange()) {
		xlsTRange rng;
		pSrc->getRange(rng);
		int rows = rng.getNrRows();
		int cols = rng.getNrCols();

		// Case 1 - cell�׷��� ũ�Ⱑ ����İ� ���ٸ�
		if (rows == c_rows && cols == c_cols) {
			pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1() + a_c);
		}
		else { // ������� ũ��� ���̳��� ��� spec�� [18.17.2.7] [Single- and Array Formulas]�� �����ȴ�� 
			// �Ķ���͵鿡 ���� ó���� �����Ѵ�.

			// ������� 1*1�����̶�� 1�� cell�� �����Ȱ�ó�� ����
			if (rows == 1 && cols == 1) {
				pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1(), pSrc->getCol1());
			}
			// Case 2 - ���� cell�׷��� ���麸�� �� ���� ����� �����ٸ� ������ �� �������(left-most columns)�� cell�鿡 �����ȴ�.
			else if (c_rows < rows && c_cols >= cols && a_r >= c_rows) {
				pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1());
			}
			// Case 3 - ���� cell�׷��� ���麸�� �� ���� �ļ��� �����ٸ� ������ �ǿ��ʷĵ�(top-most rows)�� cell�鿡 �����ȴ�.
			else if (c_cols < cols && c_rows >= rows && a_c >= c_cols) {
				pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1(), pSrc->getCol1() + a_c);
			}
			// Case 4 - ���� cell�׷��� ���麸�� ���� ����� �����ٸ� �� cell�� ������ ��츦 �����ϰ� �ڱ��� �����ġ�� �ش��� ���� ������.
			else if (c_rows >= rows && a_r >= rows) {
				// Case 4:1 - 1*N Ȥ�� 2�������� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� �ǿ�����cell���� �������� ���� ��(N/A)�� ������.
				if (a_c >= cols) {
					pDst->setError(eNA);
				}
				else if (c_rows >= 1 && rows > 1) {
					pDst->setError(eNA);
				}
				// Case 4:2 - N*1�� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� ����� ù��° ���� �����Ѵ�.
				else if (c_cols == 1) {
					pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1(), pSrc->getCol1() + a_c);
				}
				else { // ��Ÿ
					if (rows == 1) {
						pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1(), pSrc->getCol1() + a_c);
					}
					else {
						pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1() + a_c);
					}
				}
			}
			// Case 5 - ���� cell�׷��� ���麸�� ���� �ĵ��� �����ٸ� �� cell�� ������ ��츦 �����ϰ� �ڱ��� �����ġ�� �ش��� ���� ������.
			else if (c_cols >= cols && a_c >= cols) {
				// Case 5:1 - N*1 Ȥ�� 2�������� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� �ǹ��� cell���� �������� ���� ��(N/A)�� ������.
				if (a_r >= rows) {
					pDst->setError(eNA);
				}
				else if (c_cols >= 1 && cols > 1) {
					pDst->setError(eNA);
				}
				// Case 5:2 - 1*N�� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� �ĵ��� ù��° ���� �����Ѵ�.
				else if (c_rows == 1) {
					pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1());
				}
				else { // ��Ÿ
					if (cols == 1) {
						pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1());
					}
					else {
						pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1() + a_c);
					}
				}
			}
			// Case 4�� 5�� ������ 
			else {
				pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1() + a_c);
			}
		}
	}
	else if (pSrc->isArray()) {
		xlsValueArray* va = pSrc->m_array;
		int rows = va->getRowCount();
		int cols = va->getColCount();

		// Case 1 - cell�׷��� ũ�Ⱑ ����İ� ���ٸ�
		if (rows == c_rows && cols == c_cols) {
			xlsValue* v = va->getValue(a_r, a_c);
			pDst->BrCopy(v);
		}
		else { // ������� ũ��� ���̳��� ��� spec�� [18.17.2.7] [Single- and Array Formulas]�� �����ȴ�� 
			// �Ķ���͵鿡 ���� ó���� �����Ѵ�.

			// ������� 1*1�����̶�� 1�� cell�� �����Ȱ�ó�� ����
			if (rows == 1 && cols == 1) {
				xlsValue* v = va->getValue(0, 0);
				pDst->BrCopy(v);
			}
			// Case 2 - ���� cell�׷��� ���麸�� �� ���� ����� �����ٸ� ������ �� �������(left-most columns)�� cell�鿡 �����ȴ�.
			else if (c_rows < rows && c_cols >= cols && a_r >= c_rows) {
				xlsValue* v = va->getValue(a_r, 0);
				pDst->BrCopy(v);
			}
			// Case 3 - ���� cell�׷��� ���麸�� �� ���� �ļ��� �����ٸ� ������ �ǿ��ʷĵ�(top-most rows)�� cell�鿡 �����ȴ�.
			else if (c_cols < cols && c_rows >= rows && a_c >= c_cols) {
				xlsValue* v = va->getValue(0, a_c);
				pDst->BrCopy(v);
			}
			// Case 4 - ���� cell�׷��� ���麸�� ���� ����� �����ٸ� �� cell�� ������ ��츦 �����ϰ� �ڱ��� �����ġ�� �ش��� ���� ������.
			else if (c_rows >= rows && a_r >= rows) {
				// Case 4:1 - 1*N Ȥ�� 2�������� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� �ǿ�����cell���� �������� ���� ��(N/A)�� ������.
				if (a_c >= cols) {
					pDst->setError(eNA);
				}
				else if (c_rows >= 1 && rows > 1) {
					pDst->setError(eNA);
				}
				// Case 4:2 - N*1�� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� ����� ù��° ���� �����Ѵ�.
				else if (c_cols == 1) {
					xlsValue* v = va->getValue(0, a_c);
					pDst->BrCopy(v);
				}
				else { // ��Ÿ
					if (rows == 1) {
						xlsValue* v = va->getValue(0, a_c);
						pDst->BrCopy(v);
					}
					else {
						xlsValue* v = va->getValue(a_r, a_c);
						pDst->BrCopy(v);
					}
				}
			}
			// Case 5 - ���� cell�׷��� ���麸�� ���� �ĵ��� �����ٸ� �� cell�� ������ ��츦 �����ϰ� �ڱ��� �����ġ�� �ش��� ���� ������.
			else if (c_cols >= cols && a_c >= cols) {
				// Case 5:1 - N*1 Ȥ�� 2�������� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� �ǹ��� cell���� �������� ���� ��(N/A)�� ������.
				if (a_r >= rows) {
					pDst->setError(eNA);
				}
				else if (c_cols >= 1 && cols > 1) {
					pDst->setError(eNA);
				}
				// Case 5:2 - 1*N�� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� �ĵ��� ù��° ���� �����Ѵ�.
				else if (c_rows == 1) {
					xlsValue* v = va->getValue(a_r, 0);
					pDst->BrCopy(v);
				}
				else { // ��Ÿ
					if (cols == 1) {
						xlsValue* v = va->getValue(a_r, 0);
						pDst->BrCopy(v);
					}
					else {
						xlsValue* v = va->getValue(a_r, a_c);
						pDst->BrCopy(v);
					}
				}
			}
			// Case 4�� 5�� ������ 
			else {
				xlsValue* v = va->getValue(a_r, a_c);
				pDst->BrCopy(v);
			}
		}
	}
	else {
		// pDst->setError(eNAvv);
		pDst->BrCopy(pSrc);
	}
}

// [��ļ���ó��-2] 4Ģ������ ����Լ����� �־��� �Է��Ķ���Ϳ��� nIndex�� �ش��� ���� ���
// pDst : ����� ���� �����ϴ� ����
// pSrc : �Է��Ķ����
// nIndex : �Է��Ķ������ ����
// ������ : 4Ģ������ ���� token
void xlsArrayEvaluator::getValFromInputParameter(xlsCalValue* pDst, xlsCalValue* pSrc, int nIndex)
{
	//bool bContinue = false;
	// ����Ŀ� ������ �� �� �Ĺ�ȣ�� ���
	int v_rows = m_arrayResultVals->getRowCount();
	int v_cols = m_arrayResultVals->getColCount();
	int v_r = nIndex / v_cols;
	int v_c = nIndex % v_cols;

	if (pSrc->isRange()) {
		xlsTRange rng;
		pSrc->getRange(rng);
		int rows = rng.getNrRows();
		int cols = rng.getNrCols();

		// ����� ũ�Ⱑ ����İ� ���ٸ�
		if (rows == v_rows && cols == v_cols) {
			pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + v_r, pSrc->getCol1() + v_c);
		}
		else { // ������� ũ��� ���̳��ٸ�
			if (rows == 1) { // (1 * N)�� 1��������̶��
				pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1(), pSrc->getCol1() + v_c);
			}
			else if (cols == 1) { // (N * 1)�� 1��������̶��
				pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + v_r, pSrc->getCol1());
			}
			else {
				// empty process
			}
		}
	}
	else if (pSrc->isArray()) {
		xlsValueArray* va = pSrc->m_array;
		int rows = va->getRowCount();
		int cols = va->getColCount();

		// ����� ũ�Ⱑ ����İ� ���ٸ�
		if (rows == v_rows && cols == v_cols) {
			xlsValue* v = va->getValue(v_r, v_c);
			pDst->BrCopy(v);
		}
		else { // ������� ũ��� ���̳��ٸ�
			if (rows == 1) { // (1 * N)�� 1��������̶��
				xlsValue* v = va->getValue(0, v_c);
				pDst->BrCopy(v);
			}
			else if (cols == 1) { // (N * 1)�� 1��������̶��
				xlsValue* v = va->getValue(v_r, 0);
				pDst->BrCopy(v);
			}
			else {
				// empty process
			}
		}
	}
	else {
		pDst->copy(pSrc);
	}
}

// [��ļ���ó��-2] ��ļ��Ŀ� ���� xlsTokenFuncBasic��ü�� ǥ���Ǵ� �Լ�ó��(m_eXLS_TokenFuncBasic)
// xlsTokenFuncBasic�Լ��� ���� ���ó���̹Ƿ� �Է��Ķ������ ������ 1�� Ȥ�� 0����°��� ������ �Ѵ�.
// token : xlsTokenFuncBasic�Լ��� ���� token
// ������ : xlsTokenFuncBasic�Լ��� ���� token
xlsToken* xlsArrayEvaluator::processTokenFuncVarBasic(xlsToken* token)
{
	xlsToken* pToken = BrNULL;
	xlsFunc::eFuncArgs eFuncNum = (xlsFunc::eFuncArgs)token->getFuncNum();

	switch(eFuncNum) {
	case xlsFunc::eDecimal:
	case xlsFunc::eNot:
	case xlsFunc::eIsText:
	case xlsFunc::eIsNumber:
	case xlsFunc::eIsNonText:
	case xlsFunc::eIsNA:
	case xlsFunc::eIsLogical:
	case xlsFunc::eIsFormula:
	case xlsFunc::eIsError:
	case xlsFunc::eIsErr:
	case xlsFunc::eIsBlank:
	case xlsFunc::eErrorType:
	case xlsFunc::eValue:
	case xlsFunc::eUpper:
	case xlsFunc::eUniCode:
	case xlsFunc::eUniChar:
	case xlsFunc::eTrim:
	case xlsFunc::eTextFunc:
	case xlsFunc::eReplaceB:
	case xlsFunc::eReplace:
	case xlsFunc::eRept:
	case xlsFunc::eProper:
	case xlsFunc::ePhonetic:
	case xlsFunc::eMidB:
	case xlsFunc::eMid:
	case xlsFunc::eLower:
	case xlsFunc::eLenB:
	case xlsFunc::eLen:
	case xlsFunc::eExact:
	case xlsFunc::eDBCS:
	case xlsFunc::eCode:
	case xlsFunc::eClean:
	case xlsFunc::eChar:
	case xlsFunc::eAsc:
	case xlsFunc::eRoundDown:
	case xlsFunc::eRoundUp:
	case xlsFunc::eRound:
	case xlsFunc::eMod:
	case xlsFunc::eFloor:
	case xlsFunc::eCeiling:
	case xlsFunc::eTanH:
	case xlsFunc::eTan:
	case xlsFunc::eOdd:
	case xlsFunc::eLog10:
	case xlsFunc::eLn:
	case xlsFunc::eIntFunc:
	case xlsFunc::eFact:
	case xlsFunc::eEven:
	case xlsFunc::eSinH:
	case xlsFunc::eSin:
	case xlsFunc::eSecH:
	case xlsFunc::eSec:
	case xlsFunc::eCscH:
	case xlsFunc::eCsc:
	case xlsFunc::eCotH:
	case xlsFunc::eCot:
	case xlsFunc::eCosH:
	case xlsFunc::eCos:
	case xlsFunc::eATanH:
	case xlsFunc::eAtan2:
	case xlsFunc::eATan:
	case xlsFunc::eASinH:
	case xlsFunc::eASin:
	case xlsFunc::eArabic:
	case xlsFunc::eACotH:
	case xlsFunc::eACot:
	case xlsFunc::eACosH:
	case xlsFunc::eACos:
	case xlsFunc::eSLN:
	case xlsFunc::eRRI:
	case xlsFunc::ePduration:
	case xlsFunc::exBitXor:
	case xlsFunc::exBitRShift:
	case xlsFunc::exBitOr:
	case xlsFunc::exBitLShift:
	case xlsFunc::exBitAnd:
	case xlsFunc::eTimeValue:
	case xlsFunc::eTime:
	case xlsFunc::eSecond:
	case xlsFunc::eYear:
	case xlsFunc::eMonth:
	case xlsFunc::eMinute:
	case xlsFunc::eHour:
	case xlsFunc::eDays:
	case xlsFunc::eDay:
	case xlsFunc::eDateValue:
	case xlsFunc::eDate:
		pToken = doTokenFunc(token);
		break;
	case xlsFunc::eRows:
	case xlsFunc::eColumns:
	case xlsFunc::eType:
	case xlsFunc::eN:
	case xlsFunc::eIsRef:
		pToken = doTokenFunc(token, BrTRUE);
		break;
	case xlsFunc::eTrue:
	case xlsFunc::eFalse:
	case xlsFunc::eNAFunc:
	case xlsFunc::ePi:
	case xlsFunc::eRand:
	case xlsFunc::eToday:
	case xlsFunc::eNow:
		pToken = doArgs0(token);
		break;
	case xlsFunc::eT:
	case xlsFunc::eMIRR:
		pToken = doTokenFuncWithArray(token, 1);
		break;
	default:
		//		pToken = doTokenFuncBasic(token);
		break;
	}

	return pToken;
}

// ��ļ��Ŀ� ���� xlsTokenFuncBasic��ü�� ǥ���Ǵ� �Լ�ó��(m_eXLS_TokenFuncBasic)
// �� ó���ΰ� �����ϴ� ���ĵ� : (������ �˼� ����)
// token : xlsTokenFuncBasic�Լ��� ���� token
// ������ : xlsTokenFuncBasic�Լ��� ���� token
xlsToken* xlsArrayEvaluator::doTokenFuncBasic(xlsToken* token)
{
	xlsCalValue *val1 = NULL;
	xlsToken* token_backup = NULL;

	// �ʱ�ȭ
	val1 = m_evaluator->m_val;

	// ù��° �Ķ������ Backup
	m_arrayInputVal1->copy(val1);

	// ����� �����ϱ� ���� ��������
	int nRows1 = 0, nCols1 = 0;
	if (val1->isRange()) {
		xlsTRange rng;
		val1->getRange(rng);
		nRows1 = rng.getNrRows();
		nCols1 = rng.getNrCols();
	}
	else if (val1->isArray()) {
		xlsValueArray* va = val1->m_array;
		nRows1 = va->getRowCount();
		nCols1 = va->getColCount();
	}
	else {
		nRows1 = 0;
		nCols1 = 0;
	}

	if (nRows1 == 0 && nCols1 == 0)
		m_arrayResultVals->setSize(1, 1);
	else
		m_arrayResultVals->setSize(nRows1, nCols1);

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nIndex < nTotalCount) {
		getValFromInputParameter(val1, m_arrayInputVal1, nIndex);

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �缳��
		m_evaluator->m_val = val1;
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� �����(val1)�� �����ϱ�
	// �ֳ��ϸ� ��Ŀ����� ���� �Ķ���Ͱ� n���� ���(2���̻�)�� �ֱ⶧���̴�.
	int nVal = val1->m_nVal;
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// [��ļ���ó��-2] ��ļ��Ŀ� ���� xlsTokenFunc��ü�� ǥ���Ǵ� �Լ�ó��(m_eXLS_TokenFunc)
// xlsTokenFunc�Լ��� ���� ���ó���̹Ƿ� �Է��Ķ������ ������ �������̴�.
// token : xlsTokenFunc�Լ��� ���� token
// nResultCount : xlsTokenFunc�Լ������ ����
// bSingCell : ����cell�ΰ��� ��Ÿ���� ���
// ������ : xlsTokenFunc�Լ��� ���� token
xlsToken* xlsArrayEvaluator::processTokenFunc(xlsToken* token, int& nResultCount, BrBOOL bSingCell)
{
	xlsCalValue *val1 = NULL, *val2 = NULL;
	xlsToken* token_backup = NULL;

	// �ʱ�ȭ
	val1 = m_evaluator->m_val->m_prev;
	val2 = m_evaluator->m_val;
	nResultCount = 0;

	// ù��° �Ķ������ Backup
	m_arrayInputVal1->copy(val1);

	// �ι�° �Ķ������ Backup
	m_arrayInputVal2->copy(val2);

	// ����� �����ϱ� ���� ��������
	int nRows = 0, nCols = 0;
	if (val2->isRange()) {
		xlsTRange rng;
		val2->getRange(rng);
		nRows = rng.getNrRows();
		nCols = rng.getNrCols();
	}
	else if (val2->isArray()) {
		xlsValueArray* va = val2->m_array;
		nRows = va->getRowCount();
		nCols = va->getColCount();
	}
	else {
		nRows = 0;
		nCols = 0;
	}

	if (nRows == 0 && nCols == 0)
		m_arrayResultVals->setSize(1, 1);
	else
		m_arrayResultVals->setSize(nRows, nCols);

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nResultCount < nTotalCount) {
		nIndex = nResultCount;
		getValFromInputParameter(val2, m_arrayInputVal2, nIndex);

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nResultCount / m_arrayResultVals->getColCount();
		nCol = nResultCount % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nResultCount++;

		// �缳��
		m_evaluator->m_val = val2;
		m_evaluator->m_val->m_prev->copy(m_arrayInputVal1);
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� �����(val1)�� �����ϱ�
	// �ֳ��ϸ� ��Ŀ����� ���� �Ķ���Ͱ� n���� ���(2���̻�)�� �ֱ⶧���̴�.
	int nVal = val1->m_nVal;
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// ��ļ��Ŀ� ���� xlsTokenFunc��ü�� ǥ���Ǵ� �Լ�ó��(m_eXLS_TokenFunc)
xlsToken* xlsArrayEvaluator::processTokenFunc(xlsToken* token)
{
	xlsToken* pToken = BrNULL;
	xlsFunc::eFuncArgs eFuncNum = (xlsFunc::eFuncArgs)token->getFuncNum();

	switch(eFuncNum) {
	case xlsFunc::eTranspose:
	case xlsFunc::eMInverse:
		pToken = doSomeArraysFuncWithReturnArray(token, 0);
		break;
	case xlsFunc::eMMult:
	case xlsFunc::eFrequency:
		pToken = doSomeArraysFuncWithReturnArray(token, 1);
		break;
	case xlsFunc::eStandardize:
	case xlsFunc::ePhi:
	case xlsFunc::ePermutationa:
	case xlsFunc::ePermut:
	case xlsFunc::eInfo:
	case xlsFunc::eRadians:
	case xlsFunc::ePowerFunc:
	case xlsFunc::eCombinA:
	case xlsFunc::eCombin:
	case xlsFunc::eSqrt:
	case xlsFunc::eSign:
	case xlsFunc::eExpFunc:
	case xlsFunc::eDegrees:
	case xlsFunc::eAbs:
	case xlsFunc::eGauss:
	case xlsFunc::eGammaLn:
	case xlsFunc::eGamma:
	case xlsFunc::eFisherInv:
	case xlsFunc::eFisher:
	case xlsFunc::eConfidence_T:
	case xlsFunc::eConfidence_Norm:
	case xlsFunc::eChisq_Dist:
	case xlsFunc::eWeibull:
	case xlsFunc::eT_Inv_2t:
	case xlsFunc::eT_Inv:
	case xlsFunc::eTInv:
	case xlsFunc::eT_Dist_2t:
	case xlsFunc::eT_Dist_Rt:
	case xlsFunc::eT_Dist:
	case xlsFunc::eTDist:
	case xlsFunc::ePoisson:
	case xlsFunc::eNormSInv:
	case xlsFunc::eNorm_S_Dist:
	case xlsFunc::eNormSDist:
	case xlsFunc::eNormInv:
	case xlsFunc::eNormDist:
	case xlsFunc::eNegBinom_Dist:
	case xlsFunc::eNegBinomDist:
	case xlsFunc::eLogNorm_Dist:
	case xlsFunc::eLogNormDist:
	case xlsFunc::eLogInv:
	case xlsFunc::eHypGeom_Dist:
	case xlsFunc::eHypGeomDist:
	case xlsFunc::eGammaInv:
	case xlsFunc::eGammaDist:
	case xlsFunc::eF_Inv_Rt:
	case xlsFunc::eF_Inv:
	case xlsFunc::eFInv:
	case xlsFunc::eF_Dist:
	case xlsFunc::eFDist:
	case xlsFunc::eExponDist:
	case xlsFunc::eCritBinom:
	case xlsFunc::eBinomDist:
	case xlsFunc::eChiDist:
	case xlsFunc::eChisq_Inv_Left:
	case xlsFunc::eChisq_Inv_Right:
	case xlsFunc::eConfidence:
		pToken = doTokenFunc(token);
		break;
	case xlsFunc::eSumXmY2:
	case xlsFunc::eSumX2pY2:
	case xlsFunc::eSumX2mY2:
	case xlsFunc::eIntercept:
	case xlsFunc::eSteYX:
	case xlsFunc::eSlope:
	case xlsFunc::eRsq:
	case xlsFunc::ePearson:
	case xlsFunc::eAreas:
	case xlsFunc::eCorrel:
	case xlsFunc::eMDeterm:
	case xlsFunc::eFTest:
	case xlsFunc::eCovar_S:
	case xlsFunc::eCovar_P:
	case xlsFunc::eCovar:
	case xlsFunc::eChiTest:
		pToken = doTokenFunc(token, BrTRUE);
		break;
	case xlsFunc::eTrimMean:
	case xlsFunc::eQuartile_Exc:
	case xlsFunc::eQuartile:
	case xlsFunc::ePercentileExc:
	case xlsFunc::ePercentile:
	case xlsFunc::eLargeFunc:
	case xlsFunc::eSmallFunc:
		pToken = doFirstArray_Args2(token);
		break;
	case xlsFunc::eForecast:
		pToken = doTokenFuncWithArray(token, 2, BrFALSE);
		break;
	case xlsFunc::eCountBlank:
		pToken = doTokenFuncWithArray(token, 1);
		break;
	case xlsFunc::eTTest:
		pToken = doTokenFuncWithArray(token, 2);
		break;
	case xlsFunc::eDVarP:
	case xlsFunc::eDVar:
	case xlsFunc::eDSum:
	case xlsFunc::eDStDevP:
	case xlsFunc::eDStDev:
	case xlsFunc::eDProduct:
	case xlsFunc::eDMin:
	case xlsFunc::eDMax:
	case xlsFunc::eDGet:
	case xlsFunc::eDCountA:
	case xlsFunc::eDCount:
	case xlsFunc::eDAverage:
		pToken = doBothArray_Arg3(token);
		break;
	}

	return pToken;
}

// ��ļ��Ŀ� ���� xlsTokenFuncVar��ü�� ǥ���Ǵ� �Լ�ó��(m_eXLS_TokenFuncVar)
xlsToken* xlsArrayEvaluator::processTokenFuncVar(xlsToken* token)
{
	xlsToken* pToken = BrNULL;
	xlsFunc::eFuncArgs eFuncNum = (xlsFunc::eFuncArgs)token->getFuncNum();

	switch(eFuncNum) {
	case xlsFunc::eLookup:
	case xlsFunc::eSubtotal:
		pToken = doAnyOneNoneArrayFuncVar(token, 0);
		break;
	case xlsFunc::eLogEst:
	case xlsFunc::eLinEst:
		pToken = doSomeArraysFuncVarWithReturnArray(token, 1);
		break;
	case xlsFunc::eGrowth:
	case xlsFunc::eTrend:
		pToken = doSomeArraysFuncVarWithReturnArray(token, 2);
		break;
	case xlsFunc::eGetPivotData:
	case xlsFunc::eAddress:
	case xlsFunc::eIndirect:
	case xlsFunc::eSubstitute:
	case xlsFunc::eSearchB:
	case xlsFunc::eSearch:
	case xlsFunc::eRightB:
	case xlsFunc::eRight:
	case xlsFunc::eNumberValue:
	case xlsFunc::eNumberString:
	case xlsFunc::eLeftB:
	case xlsFunc::eLeft:
	case xlsFunc::eFixed:
	case xlsFunc::eFindB:
	case xlsFunc::eFind:
	case xlsFunc::eUSDollar:
	case xlsFunc::eConcatenate:
	case xlsFunc::eTrunc:
	case xlsFunc::eRoman:
	case xlsFunc::eLog:
	case xlsFunc::eFloor_Preceise:
	case xlsFunc::eFloor_Math:
	case xlsFunc::eCeiling_Math:
	case xlsFunc::eBase:
	case xlsFunc::eVDB:
	case xlsFunc::eRate:
	case xlsFunc::ePV:
	case xlsFunc::ePPmt:
	case xlsFunc::ePmt:
	case xlsFunc::eNPer:
	case xlsFunc::eIPmt:
	case xlsFunc::eFV:
	case xlsFunc::eDB:
	case xlsFunc::eDDB:
	case xlsFunc::eWeekday:
	case xlsFunc::eDays360:
	case xlsFunc::eBinomDist_range:
	case xlsFunc::eBetaInv:
	case xlsFunc::eBeta_Dist:
	case xlsFunc::eBetaDist:
		pToken = doTokenFuncVar(token);
		break;
	case xlsFunc::eRow:
		pToken = doRowColFuncVar(token, BrFALSE);
		break;
	case xlsFunc::eColumn:
		pToken = doRowColFuncVar(token, BrTRUE);
		break;
	case xlsFunc::eChoose:
	case xlsFunc::eSHEETS:
	case xlsFunc::eSHEET:
	case xlsFunc::eCountA:	
	case xlsFunc::eCount:
	case xlsFunc::eSumProduct:
	case xlsFunc::eSum:
	case xlsFunc::eXor:
	case xlsFunc::eOr:
	case xlsFunc::eAnd:
	case xlsFunc::eVarPA:
	case xlsFunc::eVarA:
	case xlsFunc::eStDevPA:
	case xlsFunc::eStDevA:
	case xlsFunc::eSkew_P:
	case xlsFunc::eMinA:
	case xlsFunc::eMin:
	case xlsFunc::eMedian:
	case xlsFunc::eMaxA:
	case xlsFunc::eMax:
	case xlsFunc::eDevSq:
	case xlsFunc::eAverageA:
	case xlsFunc::eAverage:
	case xlsFunc::eAveDev:
	case xlsFunc::eSkew:
	case xlsFunc::eKurt:
	case xlsFunc::eMode:
	case xlsFunc::eMode_Sngl:
	case xlsFunc::eGeoMean:
	case xlsFunc::eHarMean:
	case xlsFunc::eSumSq:
	case xlsFunc::eProduct:
	case xlsFunc::eVarP:
	case xlsFunc::eVar:
	case xlsFunc::eStDevP:
	case xlsFunc::eStDev:
		pToken = doTokenFuncVar(token, 1, BrTRUE);
		break;
	case xlsFunc::eHyperLink:
		pToken = doHyperlinkFuncVar(token);
		break;
	case xlsFunc::eProb:
		pToken = doSomeArraysFuncVar(token, 1);
		break;
	case xlsFunc::eCellFunc:
	case xlsFunc::eNPV:
		pToken = doTokenFuncVar(token, 2, BrTRUE);
		break;
	case xlsFunc::eZTest:
	case xlsFunc::ePercentRankExc:
	case xlsFunc::ePercentRank:
		pToken = doFirstArray_Args3_Var(token);
		break;
	case xlsFunc::eAggregate:
		pToken = doTokenFuncVar(token, 3, BrTRUE);
		break;
	case xlsFunc::eOffset:
	case xlsFunc::eIndex:
	case xlsFunc::eIRR:
		pToken = doAnyOneArrayFuncVar(token, 0);
		break;
	case xlsFunc::eRank_Eq:
	case xlsFunc::eMatch:
	case xlsFunc::eVLookup:
	case xlsFunc::eHLookup:
	case xlsFunc::eRank:
		pToken = doAnyOneArrayFuncVar(token, 1);
		break;
	case xlsFunc::eAddIn:
		pToken = processTokenAddInFunc(token);
		break;
	}

	return pToken;
}

// eAddIn�ĺ��ڸ� ������ ���ĵ��� ��ļ��� ó����(m_eXLS_TokenFuncVar)
xlsToken* xlsArrayEvaluator::processTokenAddInFunc(xlsToken* token)
{
	xlsToken* pNextToken = BrNULL;
	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	xlsCalValue* val = (*(m_evaluator->m_vals))[nVal];
	if (val->isFunc() == false) {
		return pNextToken;
	}

	// �����̸��� ���
	xlsFunc* pFunc = val->m_func;
	QString name(pFunc->m_name.data(), pFunc->m_name.size());

	// �����̸��� ���ϱ�
	if (name.compareIgnorUpperLower(QString("EDATE")) == 0 ||
		name.compareIgnorUpperLower(QString("EOMONTH")) == 0 ||
		name.compareIgnorUpperLower(QString("YEARFRAC")) == 0 ||
		// Financial
		name.compareIgnorUpperLower(QString("ACCRINT")) == 0 ||
		name.compareIgnorUpperLower(QString("ACCRINTM")) == 0 ||
		name.compareIgnorUpperLower(QString("AMORDEGRC")) == 0 ||
		name.compareIgnorUpperLower(QString("AMORLINC")) == 0 ||
		name.compareIgnorUpperLower(QString("COUPDAYBS")) == 0 ||
		name.compareIgnorUpperLower(QString("COUPDAYS")) == 0 ||
		name.compareIgnorUpperLower(QString("COUPDAYSNC")) == 0 ||
		name.compareIgnorUpperLower(QString("COUPNCD")) == 0 ||
		name.compareIgnorUpperLower(QString("COUPNUM")) == 0 ||
		name.compareIgnorUpperLower(QString("COUPPCD")) == 0 ||
		name.compareIgnorUpperLower(QString("CUMIPMT")) == 0 ||
		name.compareIgnorUpperLower(QString("CUMPRINC")) == 0 ||
		name.compareIgnorUpperLower(QString("DISC")) == 0 ||
		name.compareIgnorUpperLower(QString("DURATION")) == 0 ||
		name.compareIgnorUpperLower(QString("EFFECT")) == 0 ||
		name.compareIgnorUpperLower(QString("INTRATE")) == 0 ||
		name.compareIgnorUpperLower(QString("ISPMT")) == 0 ||
		name.compareIgnorUpperLower(QString("MDURATION")) == 0 ||
		name.compareIgnorUpperLower(QString("NOMINAL")) == 0 ||
		name.compareIgnorUpperLower(QString("ODDFPRICE")) == 0 ||
		name.compareIgnorUpperLower(QString("ODDFYIELD")) == 0 ||
		name.compareIgnorUpperLower(QString("ODDLPRICE")) == 0 ||
		name.compareIgnorUpperLower(QString("ODDLYIELD")) == 0 ||
		name.compareIgnorUpperLower(QString("PRICE")) == 0 ||
		name.compareIgnorUpperLower(QString("PRICEDISC")) == 0 ||
		name.compareIgnorUpperLower(QString("PRICEMAT")) == 0 ||
		name.compareIgnorUpperLower(QString("RECEIVED")) == 0 ||
		name.compareIgnorUpperLower(QString("TBILLEQ")) == 0 ||
		name.compareIgnorUpperLower(QString("TBILLPRICE")) == 0 ||
		name.compareIgnorUpperLower(QString("TBILLYIELD")) == 0 ||
		name.compareIgnorUpperLower(QString("YIELD")) == 0 ||
		name.compareIgnorUpperLower(QString("YIELDDISC")) == 0 ||
		name.compareIgnorUpperLower(QString("YIELDMAT")) == 0) {
			pNextToken = processTokenBondFuncVar(token);
	}
	else if (name.compareIgnorUpperLower(QString("IMPRODUCT")) == 0 ||
		name.compareIgnorUpperLower(QString("IMSUM")) == 0 ||
		name.compareIgnorUpperLower(QString("BESSELI")) == 0 ||
		name.compareIgnorUpperLower(QString("BESSELJ")) == 0 ||
		name.compareIgnorUpperLower(QString("BESSELK")) == 0 ||
		name.compareIgnorUpperLower(QString("BESSELY")) == 0 ||
		name.compareIgnorUpperLower(QString("COMPLEX")) == 0 ||
		name.compareIgnorUpperLower(QString("ERF")) == 0 ||
		name.compareIgnorUpperLower(QString("ERFC")) == 0 ||
		name.compareIgnorUpperLower(QString("ERF.PRECISE")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.ERF.PRECISE")) == 0 ||
		name.compareIgnorUpperLower(QString("ERFC.PRECISE")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.ERFC.PRECISE")) == 0 ||
		name.compareIgnorUpperLower(QString("IMABS")) == 0 ||
		name.compareIgnorUpperLower(QString("IMAGINARY")) == 0 ||
		name.compareIgnorUpperLower(QString("IMARGUMENT")) == 0 ||
		name.compareIgnorUpperLower(QString("IMCONJUGATE")) == 0 ||
		name.compareIgnorUpperLower(QString("IMCOS")) == 0 ||
		name.compareIgnorUpperLower(QString("IMCOSH")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.IMCOSH")) == 0 ||
		name.compareIgnorUpperLower(QString("IMCOT")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.IMCOT")) == 0 ||
		name.compareIgnorUpperLower(QString("IMCSC")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.IMCSC")) == 0 ||
		name.compareIgnorUpperLower(QString("IMCSCH")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.IMCSCH")) == 0 ||
		name.compareIgnorUpperLower(QString("IMDIV")) == 0 ||
		name.compareIgnorUpperLower(QString("IMEXP")) == 0 ||
		name.compareIgnorUpperLower(QString("IMLN")) == 0 ||
		name.compareIgnorUpperLower(QString("IMLOG10")) == 0 ||
		name.compareIgnorUpperLower(QString("IMLOG2")) == 0 ||
		name.compareIgnorUpperLower(QString("IMPOWER")) == 0 ||
		name.compareIgnorUpperLower(QString("IMREAL")) == 0 ||
		name.compareIgnorUpperLower(QString("IMSEC")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.IMSEC")) == 0 ||
		name.compareIgnorUpperLower(QString("IMSECH")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.IMSECH")) == 0 ||
		name.compareIgnorUpperLower(QString("IMSIN")) == 0 ||
		name.compareIgnorUpperLower(QString("IMSINH")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.IMSINH")) == 0 ||
		name.compareIgnorUpperLower(QString("IMSQRT")) == 0 ||
		name.compareIgnorUpperLower(QString("IMSUB")) == 0 ||
		name.compareIgnorUpperLower(QString("IMTAN")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.IMTAN")) == 0) {
			pNextToken = processTokenEngineeringFuncVar(token);
	}
	else if (name.compareIgnorUpperLower(QString("ISOWEEKNUM")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.ISOWEEKNUM")) == 0 ||
		name.compareIgnorUpperLower(QString("WORKDAY")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.WORKDAY")) == 0 ||
		name.compareIgnorUpperLower(QString("WORKDAY.INTL")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.WORKDAY.INTL")) == 0 ||
		name.compareIgnorUpperLower(QString("NETWORKDAYS")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.NETWORKDAYS")) == 0 ||
		name.compareIgnorUpperLower(QString("NETWORKDAYS.INTL")) == 0 ||
		name.compareIgnorUpperLower(QString("_XLFN.NETWORKDAYS.INTL")) == 0 ||
		name.compareIgnorUpperLower(QString("WEEKNUM")) == 0 ||
		name.compareIgnorUpperLower(QString("BIN2DEC")) == 0 ||
		name.compareIgnorUpperLower(QString("BIN2HEX")) == 0 ||
		name.compareIgnorUpperLower(QString("BIN2OCT")) == 0 ||
		name.compareIgnorUpperLower(QString("DEC2BIN")) == 0 ||
		name.compareIgnorUpperLower(QString("DEC2HEX")) == 0 ||
		name.compareIgnorUpperLower(QString("DEC2OCT")) == 0 ||
		name.compareIgnorUpperLower(QString("DELTA")) == 0 ||
		name.compareIgnorUpperLower(QString("GESTEP")) == 0 ||
		name.compareIgnorUpperLower(QString("HEX2BIN")) == 0 ||
		name.compareIgnorUpperLower(QString("HEX2DEC")) == 0 ||
		name.compareIgnorUpperLower(QString("HEX2OCT")) == 0 ||
		name.compareIgnorUpperLower(QString("OCT2BIN")) == 0 ||
		name.compareIgnorUpperLower(QString("OCT2DEC")) == 0 ||
		name.compareIgnorUpperLower(QString("OCT2HEX")) == 0 ||
		name.compareIgnorUpperLower(QString("DOLLARDE")) == 0 ||
		name.compareIgnorUpperLower(QString("DOLLARFR")) == 0 ||
		name.compareIgnorUpperLower(QString("FVSCHEDULE")) == 0 || 
		name.compareIgnorUpperLower(QString("XIRR")) == 0 ||
		name.compareIgnorUpperLower(QString("XNPV")) == 0 ||
		name.compareIgnorUpperLower(QString("FACTDOUBLE")) == 0 ||
		name.compareIgnorUpperLower(QString("SQRTPI")) == 0 ||
		name.compareIgnorUpperLower(QString("GCD")) == 0 ||
		name.compareIgnorUpperLower(QString("LCM")) == 0 ||
		name.compareIgnorUpperLower(QString("MROUND")) == 0 ||
		name.compareIgnorUpperLower(QString("MULTINOMIAL")) == 0 ||
		name.compareIgnorUpperLower(QString("QUOTIENT")) == 0 ||
		name.compareIgnorUpperLower(QString("RANDBETWEEN")) == 0 ||
		name.compareIgnorUpperLower(QString("SERIESSUM")) == 0 ||
		name.compareIgnorUpperLower(QString("ISEVEN")) == 0 ||
		name.compareIgnorUpperLower(QString("ISODD")) == 0 ||
		name.compareIgnorUpperLower(QString("CONVERT")) == 0) {
			pNextToken = processTokenMiscAddinFuncVar(token);
	}
	else {
		// �� ó��
	}

	return pNextToken;
}

// xlsEngineerFuncsŬ�󽺷� ǥ���Ǵ� ���ĵ鿡 ���� ��ļ���ó����
// �� �Լ����� ó���ϴ� ���ĵ��� �Ƿ� : IMPRODUCT, IMSUM ��
xlsToken* xlsArrayEvaluator::processTokenEngineeringFuncVar(xlsToken* token)
{
	xlsToken* pNextToken = BrNULL;

	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	xlsCalValue* val = (*(m_evaluator->m_vals))[nVal];
	if (val->isFunc() == false) {
		return pNextToken;
	}

	// [��ļ��ı���-3�ܰ�] ��������� �Ķ���͵��� �������� �ʴ� ���ĵ��� ó��
	xlsEngineerFuncs* func = (xlsEngineerFuncs*)val->m_func;
	ENUM_SUPPORT_RESULT result = checkSupportArrayParameter(token, ADDIN_ENGINEERING, func->m_nID);
	if (result == UNSUPPORT_PROCESSED) {
		pNextToken = token->m_next;
		return pNextToken;
	}

	switch(func->m_nID) {
	case eImtan:
	case eImsub:
	case eImsqrt:
	case eImsinh:
	case eImsin:
	case eImsech:
	case eImsec:
	case eImreal:
	case eImpower:
	case eImlog2:
	case eImlog10:
	case eImln:
	case eImexp:
	case eImdiv:
	case eImcsch:
	case eImcsc:
	case eImcot:
	case eImcosh:
	case eImcos:
	case eImconjugate:
	case eImargument:
	case eImaginary:
	case eImabs:
	case eErfc_Precise:
	case eErf_Precise:
	case eErfc:
	case eErf:
	case eComplex:
	case eBesselY:
	case eBesselK:
	case eBesselJ:
	case eBesselI:
		pNextToken = doTokenFuncVar(token);
		break;
	case eImproduct: // �Լ������� ������ ����.
	case eImsum:
		pNextToken = doTokenFuncVar(token, 1, BrTRUE);
		break;
	}

	return pNextToken;
}

// xlsEngineerFuncsŬ�󽺷� ǥ���Ǵ� ���ĵ鿡 ���� ��ļ���ó����
// �� �Լ����� ó���ϴ� ���ĵ��� �Ƿ� : EDATE, EOMONTH ��
xlsToken* xlsArrayEvaluator::processTokenBondFuncVar(xlsToken* token)
{
	xlsToken* pNextToken = BrNULL;

	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	xlsCalValue* val = (*(m_evaluator->m_vals))[nVal];
	if (val->isFunc() == false) {
		return pNextToken;
	}

	// [��ļ��ı���-3�ܰ�] ��������� �Ķ���͵��� �������� �ʴ� ���ĵ��� ó��
	xlsBondFuncs* func = (xlsBondFuncs*)val->m_func;
	ENUM_SUPPORT_RESULT result = checkSupportArrayParameter(token, ADDIN_BOND, func->m_nID);
	if (result == UNSUPPORT_PROCESSED) {
		pNextToken = token->m_next;
		return pNextToken;
	}

	// �Ϲ�ó��
	switch(func->m_nID) {
	case eYieldmat:
	case eYielddisc:
	case eYield:
	case eTbillyield:
	case eTbillprice:
	case eTbilleq:
	case eReceived:
	case ePricemat:
	case ePricedisc:
	case ePrice:
	case eOddlyield:
	case eOddlprice:
	case eOddfyield:
	case eOddfprice:
	case eNominal:
	case eMduration:
	case eIspmt:
	case eIntrate:
	case eEffect:
	case eDisc:
	case eDuration:
	case eCumprinc:
	case eCumipmt:
	case eCouppcd:
	case eCoupnum:
	case eCoupncd:
	case eCoupdaysnc:
	case eCoupdays:
	case eCoupdaysbs:
	case eAmorlinc:
	case eAmordegrc:
	case eAccrintm:
	case eAccrint:
	case eYearFrac:
	case eEomonth:
	case eEdate:
		pNextToken = doTokenFuncVar(token);
		break;
	}

	return pNextToken;
}

// xlsMiscAddinFuncsŬ�󽺷� ǥ���Ǵ� ���ĵ鿡 ���� ��ļ���ó����
// �� �Լ����� ó���ϴ� ���ĵ��� �Ƿ� : ISOWEEKNUM, WORKDAY.INTL ��
xlsToken* xlsArrayEvaluator::processTokenMiscAddinFuncVar(xlsToken* token)
{
	xlsToken* pNextToken = BrNULL;

	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	xlsCalValue* val = (*(m_evaluator->m_vals))[nVal];
	if (val->isFunc() == false) {
		return pNextToken;
	}

	// [��ļ��ı���-3�ܰ�] ��������� �Ķ���͵��� �������� �ʴ� ���ĵ��� ó��
	xlsMiscAddinFuncs* func = (xlsMiscAddinFuncs*)val->m_func;
	ENUM_SUPPORT_RESULT result = checkSupportArrayParameter(token, ADDIN_MISC, func->m_nID);
	if (result == UNSUPPORT_PROCESSED) {
		pNextToken = token->m_next;
		return pNextToken;
	}

	switch(func->m_nID) {
	case eSeriesSum:
		pNextToken = doTokenFuncVar(token, 5, BrTRUE);
		break;
	case eMultinomial:
	case eLCM:
	case eGCD:
		pNextToken = doTokenFuncVar(token, 1, BrTRUE);
		break;
	case eXnpv:
		pNextToken = doTokenFuncVar(token, 3, BrTRUE);
		break;
	case eXirr:
		pNextToken = doSomeArraysFuncVar(token, 2);
		break;
	case eFVSchedule:
		pNextToken = doAnyOneArrayFuncVar(token, 2);
		break;
	case eNetWorkDays_Intl:
	case eWorkday_Intl:
		pNextToken = doWorkDaySerialFuncVar(token);
		break;
	case eNetWorkDays:
	case eWorkday:
		pNextToken = doWorkDaySerialFuncVar(token, BrTRUE);
		break;
	case eConvert:
	case eIsOdd:
	case eIsEven:
	case eRandbetween:
	case eQuotient:
	case eMRound:
	case eSqrtPI:
	case eFactDouble:
	case eDollarfr:
	case eDollarde:
	case eOct2hex:
	case eOct2dec:
	case eOct2bin:
	case eHex2oct:
	case eHex2dec:
	case eHex2bin:
	case eGEStep:
	case eDelta:
	case eDec2oct:
	case eDec2hex:
	case eDec2bin:
	case eBin2oct:
	case eBin2hex:
	case eBin2dec:
	case eWeeknum:
	case eIsoWeekNum:
		pNextToken = doTokenFuncVar(token);
		break;
	}

	return pNextToken;
}

///////////////////////////////////////////////////////////////////////////
/**************** Array Formulaó���� ���� �⺻���ĵ��� hanlder ************/
///////////////////////////////////////////////////////////////////////////

// �Ķ���Ͱ����� ������ �Ϲݰ��ĵ鿡 ���� ó����(����ó����)
// ��ǥ���� ���� : CONFIDENCE
// token : �ش� ���Ŀ� ���� token�ڷ�
// bIsAllArray : ��� �Է��Ķ���͵��� spec�� ���� array�ΰ��� ��Ÿ���� ��ߺ���(�������� FALSE)
//				CHITEST, CHISQ.TEST���� �Ϻ� ���ĵ��� �Է��Ķ���ͷ� ����� �䱸�Ѵ�.
xlsToken* xlsArrayEvaluator::doTokenFunc(xlsToken* token, BrBOOL bIsAllArray)
{
	// �Լ��� �Ķ���������� ���
	xlsTokenFunc* pToken = (xlsTokenFunc*)token;

	int nArgCount = (int)pToken->getFunc()->m_nMinArgs;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0; // ��ȯ����
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			if (bIsAllArray == BrTRUE)
				vals[i]->copy(m_arrayInputVals[i]);
			else
				getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
		}

		token_backup = token->evaluate(m_evaluator);
		// BITAND�� ���� �Ϻ� ���ĵ鿡���� ������� ������ (*m_evaluator->m_vals)�� ���ΰ�(0)�� 
		// m_evaluator->m_val�� ���ΰ�(1)�� ��ġ���� �ʴ´�.
		// ���� �̿� ���� ������ �ʿ��ϴ�.(void xlsFunc::evaluate(xlsEvaluator* eval)�Լ��� �����Ұ�)
		if (m_evaluator->m_val->m_nVal != nVal) {
			m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
		}

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �缳��
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// �Ķ���Ͱ����� �������� �Ϲݰ��ĵ鿡 ���� ó����(����ó����)
// ��ǥ���� ���� : BETADIST
// nFixParaCount : �����Ǵ� �Ķ���Ͱ����� ��Ÿ����. bIsAllRefArgs=TRUE�϶����� ��ȿ�ϴ�.
//				�Ƿʷ� STDEV�� ���ؼ��� nFixParaCount = 1, NPV�� ���ؼ��� nFixParaCount = 2�� �ȴ�.
// bIsAllRefArgs : �Է��Ķ���͵��� �������� �ڷ�����(���, ����, �� ��)�� ������ �����ɼ� �ִ�
//				1~255�������� �������� Ư���� �����°��� ��Ÿ���� ��ߺ���(�������� FALSE)
//				TRUE�� �����ϸ� �̿� ���� Ư���� ������ ���ĵ�(�Ƿʷ� STDEV)�� ó���Ѵ�.
xlsToken* xlsArrayEvaluator::doTokenFuncVar(xlsToken* token, BrINT nFixParaCount, BrBOOL bIsAllRefArgs)
{
	// [��ļ��ı���-3�ܰ�] ����Ķ���͹����� �����ΰ��� �Ǵ��ϱ�
	xlsToken* token_backup = BrNULL;
	BrBOOL bSupported = BrTRUE;
	ENUM_SUPPORT_RESULT result = checkSupportArrayParameter(token);
	if (result == UNSUPPORT_PROCESSED) { // �̹� ó���ǿ��ٸ�
		token_backup = token->m_next;
		return token_backup;
	}
	else {
		if (result == SUPPORTED)
			bSupported = BrTRUE; // �����ϴ� ����
		else
			bSupported = BrFALSE; // �������� �ʴ� ����
	}

	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0; // ��ȯ����
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
	while(nIndex < nTotalCount) {
		if (bIsAllRefArgs == BrTRUE) {
			for (int i = 0; i < nArgCount; i++) {
				// AddIn���� Ư���� ������ ��Ÿ���� ������ �Ķ���ͷĿ� ���Եǿ� ���´�.
				if (token->getFuncNum() == xlsFunc::eAddIn && m_arrayInputVals[i]->isFunc())
					vals[i]->copy(m_arrayInputVals[i]);
				else if (i < (nFixParaCount - 1))
					getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
				else
					vals[i]->copy(m_arrayInputVals[i]);
			}
		}
		else {
			for (int i = 0; i < nArgCount; i++) {
				// AddIn���� Ư���� ������ ��Ÿ���� ������ �Ķ���ͷĿ� ���Եǿ� ���´�.
				if (token->getFuncNum() == xlsFunc::eAddIn && m_arrayInputVals[i]->isFunc()) {
					vals[i]->copy(m_arrayInputVals[i]);
				}
				else {
					getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
				}
			}
		}

		token_backup = token->evaluate(m_evaluator);

		// [��ļ��ı���-3�ܰ�] ��������� �Ķ���͵��� �������� �ʴ� ���ĵ��� ó��
		// ���⼭ ó���Ǵ� ���ĵ��� ��ļ������������ ����Ķ������ ũ�⸦ �Ѿ�� ���
		// �Ѿ�� ������ �κ��� #NA�� ó���ǰ� �ؾ� �� ���ĵ鿡 ���ؼ��� ����ȴ�.
		if (!bSupported) {
			if (m_evaluator->m_val->isNA() == false)
				m_evaluator->m_val->setError(eInvalidValue);
		}
		else {
			// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
			if (m_evaluator->m_val->isCell()) {
				m_evaluator->m_val->checkValue(m_evaluator);
			}
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �Ķ������ �缳��
		// SHEETS�� ���� �Ķ���Ͱ����� 0~1��(����)�� ��쵵 �����ϹǷ�...
		if (nArgCount > 0) {
			m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
		}
		else {
			m_evaluator->m_val = (*m_evaluator->m_vals)[0];
		}
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// �Ķ���Ͱ� 2���̰� ù��° �Ķ���ʹ� Array(Ȥ�� Range)�� ���ĵ鿡 ���� ó����
// �� ó���ΰ� �����ϴ� ���ĵ� : SMALL, LARGE, PERCENTILE ���
xlsToken* xlsArrayEvaluator::doFirstArray_Args2(xlsToken* token)
{
	xlsCalValue *val1 = NULL, *val2 = NULL;
	xlsToken* token_backup = NULL;

	// �ʱ�ȭ
	val1 = m_evaluator->m_val->m_prev;
	val2 = m_evaluator->m_val;

	// ù��° �Ķ������ Backup
	m_arrayInputVal1->copy(val1);

	// �ι�° �Ķ������ Backup
	m_arrayInputVal2->copy(val2);

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nIndex < nTotalCount) {
		getValInFunc(val2, m_arrayInputVal2, nIndex);

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �缳��
		m_evaluator->m_val = val2;
		m_evaluator->m_val->m_prev->copy(m_arrayInputVal1);
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� �����(val1)�� �����ϱ�
	// �ֳ��ϸ� ��Ŀ����� ���� �Ķ���Ͱ� n���� ���(2���̻�)�� �ֱ⶧���̴�.
	int nVal = val1->m_nVal;
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// �Ķ���Ͱ� 3���̰� ù��° �Ķ���ʹ� Array(Ȥ�� Range), 3��° �Ķ���ʹ� optional�� ���ĵ鿡 ���� ó����
// �� ó���ΰ� �����ϴ� ���ĵ� : PERCENTRANK, PERCENTRANK.EXC, PERCENTRANK.INC
xlsToken* xlsArrayEvaluator::doFirstArray_Args3_Var(xlsToken* token)
{
	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	xlsCalValue *val1 = BrNULL, *val2 = BrNULL, *val3 = BrNULL;
	xlsToken* token_backup = BrNULL;

	// �ʱ�ȭ
	val1 = (*m_evaluator->m_vals)[nVal];
	val2 = (*m_evaluator->m_vals)[nVal + 1];

	// ù��° �Ķ������ Backup
	m_arrayInputVal1->copy(val1);

	// �ι�° �Ķ������ Backup
	m_arrayInputVal2->copy(val2);

	// ����° �Ķ������ Backup
	if (nArgCount > 2) {
		val3 = (*m_evaluator->m_vals)[nVal + 2];
		m_arrayInputVal3->copy(val3);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nIndex < nTotalCount) {
		getValInFunc(val2, m_arrayInputVal2, nIndex);
		if (nArgCount > 2) {
			getValInFunc(val3, m_arrayInputVal3, nIndex);
		}

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �缳��
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
		// ù��° �Ķ���͸� �״�� ����
		(*m_evaluator->m_vals)[nVal]->copy(m_arrayInputVal1);
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� �����(val1)�� �����ϱ�
	// �ֳ��ϸ� ��Ŀ����� ���� �Ķ���Ͱ� n���� ���(2���̻�)�� �ֱ⶧���̴�.
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// �Ķ���Ͱ� n���� �����̰� �� Ȥ�� ���� � �Ķ���ʹ� Array(Ȥ�� Range)�� ���ĵ鿡 ���� ó����
// �� ó���ΰ� �����ϴ� �Ƿ� ���ĵ� : TTEST, T.TEST(4������, ù��°, �ι�°�� ���)
// nArrayParameterCount : Array�� �Ķ���Ͱ���
// bForward : Array�� �Ķ���͵��� �Ķ���ͷ��� ���ʿ� �ִ��� Ȥ�� ���ʿ� �ִ��� ��Ÿ���� ��ߺ���
//			true : ���ʿ� ����. false : �ڿ� ����.
xlsToken* xlsArrayEvaluator::doTokenFuncWithArray(xlsToken* token, BrINT nArrayParameterCount, BrBOOL bForward)
{
	// �Լ��� �Ķ���������� ���
	xlsTokenFunc* pToken = (xlsTokenFunc*)token;

	int nArgCount = (int)pToken->getFunc()->m_nMinArgs; 
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0; // ��ȯ����
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		if (bForward) {
			for (int i = 0; i < nArgCount; i++) {
				if (i < nArrayParameterCount)
					vals[i]->copy(m_arrayInputVals[i]);
				else
					getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
			}
		}
		else {
			for (int i = 0; i < nArgCount; i++) {
				if (i < (nArgCount - nArrayParameterCount))
					getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
				else
					vals[i]->copy(m_arrayInputVals[i]);
			}
		}

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �缳��
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// �Ķ���Ͱ� 3���� �����̰� ù��° �� ����° �Ķ���͵��� Array(Ȥ�� Range)�� ���ĵ鿡 ���� ó����
// �� ó���ΰ� �����ϴ� �Ƿ� ���ĵ� : DAVERAGE, DCOUNT (��ü�� �ڷ�������� �Լ���)
xlsToken* xlsArrayEvaluator::doBothArray_Arg3(xlsToken* token)
{
	xlsCalValue *val1 = BrNULL, *val2 = BrNULL, *val3 = BrNULL;
	xlsToken* token_backup = BrNULL;

	// �Լ��� �Ķ���������� ���
	xlsTokenFunc* pToken = (xlsTokenFunc*)token;

	int nArgCount = (int)pToken->getFunc()->m_nMinArgs; 
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �ʱ�ȭ
	val1 = (*m_evaluator->m_vals)[nVal + 0];
	val2 = (*m_evaluator->m_vals)[nVal + 1];
	val3 = (*m_evaluator->m_vals)[nVal + 2];

	// ù��° �Ķ������ Backup
	m_arrayInputVal1->copy(val1);

	// �ι�° �Ķ������ Backup
	m_arrayInputVal2->copy(val2);

	// ����° �Ķ������ Backup
	m_arrayInputVal3->copy(val3);

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nIndex < nTotalCount) {
		val1->copy(m_arrayInputVal1);
		getValInFunc(val2, m_arrayInputVal2, nIndex);
		val3->copy(m_arrayInputVal3);

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �缳��
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� �����(val1)�� �����ϱ�
	// �ֳ��ϸ� ��Ŀ����� ���� �Ķ���Ͱ� n���� ���(2���̻�)�� �ֱ⶧���̴�.
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// �Ķ���Ͱ� 2���� �����̰� 2���� optional�̸� 2���� optional �Ķ������
// �Ѱ��� �Ķ���Ͱ� Array(Ȥ�� Range)�� ���ĵ鿡 ���� ó����
// �� ó���ΰ� �����ϴ� �Ƿ� ���ĵ� : NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL
// bThirdArray : ����° �Ķ���� Ȥ�� �׹�° �Ķ���Ͱ� Array�ΰ��� ��Ÿ���� ��ߺ���
//				TRUE : ����° �Ķ���Ͱ� array, FALSE : �׹�° �Ķ���Ͱ� array
xlsToken* xlsArrayEvaluator::doWorkDaySerialFuncVar(xlsToken* token, BrBOOL bThirdArray)
{
	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�
	xlsCalValue* val = (*(m_evaluator->m_vals))[nVal];

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0; // ��ȯ����
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		BrBOOL bArrayFound = BrFALSE;

		for (int i = 0; i < nArgCount; i++) {
			// AddIn���� Ư���� ������ ��Ÿ���� ������ �Ķ���ͷĿ� ���Եǿ� ���´�.
			if (token->getFuncNum() == xlsFunc::eAddIn && m_arrayInputVals[i]->isFunc()) {
				vals[i]->copy(m_arrayInputVals[i]);
			}
			else if (i == 3 && bThirdArray) {
				vals[i]->copy(m_arrayInputVals[i]);
			}
			else if (i == 4 && !bThirdArray) {
				vals[i]->copy(m_arrayInputVals[i]);
			}
			else {
				getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
			}

			// bThirdArray������ �����Ͽ� eNetWorkDays_Intl, eWorkday_Intl�� ���ؼ��� 
			// ��� �Ķ���͵��� ����ΰ��� �˻��ϱ�
			if (!bThirdArray && !bArrayFound && i != 4) {
				bArrayFound = checkArrayValue(m_arrayInputVals[i]);
			}
		}

		token_backup = token->evaluate(m_evaluator);

		// [��ļ��ı���-3�ܰ�] ��������� �Ķ���͵��� �������� �ʴ� ���ĵ��� ó��
		// ���⼭ ó���Ǵ� ���ĵ��� ��ļ������������ ����Ķ������ ũ�⸦ �Ѿ�� ���
		// �Ѿ�� ������ �κ��� #NA�� ó���ǰ� �ؾ� �� ���ĵ鿡 ���ؼ��� ����ȴ�.
		if (!bThirdArray && bArrayFound) {
			if (m_evaluator->m_val->isNA() == false)
				m_evaluator->m_val->setError(eInvalidValue);
		}
		else {
			// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
			if (m_evaluator->m_val->isCell()) {
				m_evaluator->m_val->checkValue(m_evaluator);
			}
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �Ķ������ �缳��
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// �Ķ���Ͱ� 0���� ���ĵ鿡 ���� ó����
// �� ó���ΰ� �����ϴ� �Ƿ� ���ĵ� : NOW, TODAY, RAND
xlsToken* xlsArrayEvaluator::doArgs0(xlsToken* token)
{
	int nVal = 0;

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0; // ��ȯ����
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// �Ķ���Ͱ� �����̰� ���� �Ѱ��� �Ķ���ʹ� Array(Ȥ�� Range)�� ���ĵ鿡 ���� ó����
// �� ó���ΰ� �����ϴ� ���ĵ� : FVSCHEDULE, RANK, RANK.AVG, RANK.EQ
// nArrayIndex : ����� �Ķ������ ���ΰ�(0 based index)
xlsToken* xlsArrayEvaluator::doAnyOneArrayFuncVar(xlsToken* token, BrINT nArrayIndex)
{
	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0; // ��ȯ����
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			// AddIn���� Ư���� ������ ��Ÿ���� ������ �Ķ���ͷĿ� ���Եǿ� ���´�.
			if (token->getFuncNum() == xlsFunc::eAddIn && m_arrayInputVals[i]->isFunc()) {
				vals[i]->copy(m_arrayInputVals[i]);
			}
			else if (i == nArrayIndex) {
				vals[i]->copy(m_arrayInputVals[i]);
			}
			else {
				getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
			}
		}

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �Ķ������ �缳��
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// �Ķ���Ͱ� �����̰� ���� ���� � �Ķ���ʹ� Array(Ȥ�� Range)�� ���ĵ鿡 ���� ó����
// �� ó���ΰ� �����ϴ� ���ĵ� : XIRR
// nLastArrayParaIndex : ����� �Ķ���͵��� ������ �Ķ������ ����(0 based index)
xlsToken* xlsArrayEvaluator::doSomeArraysFuncVar(xlsToken* token, BrINT nLastArrayParaIndex)
{
	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0; // ��ȯ����
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			// AddIn���� Ư���� ������ ��Ÿ���� ������ �Ķ���ͷĿ� ���Եǿ� ���´�.
			if (token->getFuncNum() == xlsFunc::eAddIn && m_arrayInputVals[i]->isFunc()) {
				vals[i]->copy(m_arrayInputVals[i]);
			}
			else if (i <= nLastArrayParaIndex) {
				vals[i]->copy(m_arrayInputVals[i]);
			}
			else {
				getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
			}
		}

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �Ķ������ �缳��
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// COLMUN, ROW���ĵ鿡 ���� ó����
// �� ���ĵ��� �Ķ���Ͱ� 0�� ��� ��İ��ĵ��� ����Ǵ� �ش� cell���� ��ġ�� �����ǹǷ�
// �����Լ��� ���� �ʿ�ó���� �ʿ��ϴ�.
xlsToken* xlsArrayEvaluator::doRowColFuncVar(xlsToken* token, BrBOOL bIsColumn)
{
	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0; // ��ȯ����
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();

		for (int i = 0; i < nArgCount; i++) {
			int m = 0;

			// COLUMN�� ��� �Է��Ķ������ ũ�⿡ ������� column�������� �����Ѵ�.
			if (bIsColumn) {
				m = nCol;
			}
			else { // ROW�� ��� �Է��Ķ������ ũ�⿡ ������� column�������� �����Ѵ�.
				m = nRow * m_arrayResultVals->getColCount();

			}
			getValInFunc(vals[i], m_arrayInputVals[i], m);
		}

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		if (nArgCount > 0) {
			m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);
		}
		else {
			int nValue = m_evaluator->m_val->getNumber();
			if (bIsColumn) { // COLUMN�����̶��
				nValue = nValue + nCol;
			}
			else { // ROW�����̶��
				nValue = nValue + nRow;
			}

			m_arrayResultVals->getValue(nRow, nCol)->setValue(nValue);
		}

		nIndex++;

		// �Ķ������ �缳��
		if (nArgCount > 0) {
			m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
		}
		else {
			m_evaluator->m_val = (*m_evaluator->m_vals)[0];
		}
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// HYPERLINK���Ŀ� ���� ó����
// �ι�° �Ķ���Ͱ� Range�� �Ǵ� ��� Range�� (0, 0)�� �ش��� ������ �����Ѵ�.(�ߺ�ȭ���� ����)
// ���� �̿� ���Ͽ� �����Լ��� ���� �ʿ�ó���� �ʿ��ϴ�.
xlsToken* xlsArrayEvaluator::doHyperlinkFuncVar(xlsToken* token)
{
	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0; // ��ȯ����
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			if (i < 1)
				getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
			else
				getValInFunc(vals[i], m_arrayInputVals[i], 0);
		}

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �Ķ������ �缳��
		// SHEETS�� ���� �Ķ���Ͱ����� 0~1��(����)�� ��쵵 �����ϹǷ�...
		if (nArgCount > 0) {
			m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
		}
		else {
			m_evaluator->m_val = (*m_evaluator->m_vals)[0];
		}
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// �Ķ���Ͱ� �����̰� ���� ���� � �Ķ���ʹ� Array(Ȥ�� Range)�̰� �������� ����� ���ĵ鿡 ���� ó����
// �� ó���ΰ� �����ϴ� ���ĵ� : TREND, LINEST, LOGEST, GROWTH
// nLastArrayParaIndex : ����� �Ķ���͵��� ������ �Ķ������ ����(0 based index)
xlsToken* xlsArrayEvaluator::doSomeArraysFuncVarWithReturnArray(xlsToken* token, BrINT nLastArrayParaIndex)
{
	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nRow = m_arrayResultVals->getRowCount();
	int nCol = m_arrayResultVals->getColCount();
	xlsToken* token_backup = BrNULL;

	// ����� �����ϴ� ��ĺ����� NA�� �ʱ�ȭ�ϱ�
	for (int r = 0; r < nRow; r++) {
		for (int c = 0; c < nCol; c++) {
			m_arrayResultVals->getValue(r, c)->setError(eNA);
		}
	}

	// �Ķ���͵��� �����ϱ�
	for (int i = 0; i < nArgCount; i++) {
		// AddIn���� Ư���� ������ ��Ÿ���� ������ �Ķ���ͷĿ� ���Եǿ� ���´�.
		if (token->getFuncNum() == xlsFunc::eAddIn && m_arrayInputVals[i]->isFunc()) {
			vals[i]->copy(m_arrayInputVals[i]);
		}
		else if (i <= nLastArrayParaIndex) {
			vals[i]->copy(m_arrayInputVals[i]);
		}
		else {
			// ����Ķ���Ͱ� �ƴ� �Ķ���͵��� ��ķ� �Ѿ�� ������ 0��° ��Ҹ��� �����ϰ� �Ѵ�.
			if (m_arrayInputVals[i]->isRange() || m_arrayInputVals[i]->isArray())
				getValInFunc(vals[i], m_arrayInputVals[i], 0);
			else
				vals[i]->copy(m_arrayInputVals[i]);
		}
	}

	// ������ ����ϱ�
	token_backup = token->evaluate(m_evaluator);

	// ������� �����ϱ�
	if (m_evaluator->m_val->m_array->getValue(0, 0)->isError()) {
		// �������� ������ �߻��Ͽ��ٸ� ����� 0��° ��ҿ� ��������
		// �����ǿ������Ƿ� �װ��� �������⿡ �����Ѵ�.
		QValueArray* srcRow = m_evaluator->m_val->m_array->getRow(0);
		for (int r = 0; r < nRow; r++) {
			QValueArray* dstRow = m_arrayResultVals->getRow(r);
			for (int c = 0; c < nCol; c++) {
				(*dstRow)[c]->BrCopy((*srcRow)[0]);
			}
		}
	}
	else {
		int nIndex = 0; // ��ȯ����
		int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
		while(nIndex < nTotalCount) {
			nRow = nIndex / m_arrayResultVals->getColCount();
			nCol = nIndex % m_arrayResultVals->getColCount();
			xlsValue* pValue = m_arrayResultVals->getValue(nRow, nCol);
			getResultValInFunc(pValue, m_evaluator->m_val->m_array, nIndex);

			nIndex++;
		}
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// �Ķ���Ͱ� �����̰� ���� ���� � �Ķ���ʹ� Array(Ȥ�� Range)�̰� �������� ����� ���ĵ鿡 ���� ó����
// �� ó���ΰ� �����ϴ� ���ĵ� : FREQUENCY
// nLastArrayParaIndex : ����� �Ķ���͵��� ������ �Ķ������ ����(0 based index)
xlsToken* xlsArrayEvaluator::doSomeArraysFuncWithReturnArray(xlsToken* token, BrINT nLastArrayParaIndex)
{
	// �Լ��� �Ķ���������� ���
	xlsTokenFunc* pToken = (xlsTokenFunc*)token;
	int nArgCount = (int)pToken->getFunc()->m_nMinArgs;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nRow = m_arrayResultVals->getRowCount();
	int nCol = m_arrayResultVals->getColCount();
	xlsToken* token_backup = BrNULL;

	// ����� �����ϴ� ��ĺ����� NA�� �ʱ�ȭ�ϱ�
	for (int r = 0; r < nRow; r++) {
		for (int c = 0; c < nCol; c++) {
			m_arrayResultVals->getValue(r, c)->setError(eNA);
		}
	}

	// �Ķ���͵��� �����ϱ�
	for (int i = 0; i < nArgCount; i++) {
		// AddIn���� Ư���� ������ ��Ÿ���� ������ �Ķ���ͷĿ� ���Եǿ� ���´�.
		if (token->getFuncNum() == xlsFunc::eAddIn && m_arrayInputVals[i]->isFunc()) {
			vals[i]->copy(m_arrayInputVals[i]);
		}
		else if (i <= nLastArrayParaIndex) {
			vals[i]->copy(m_arrayInputVals[i]);
		}
		else {
			// ����Ķ���Ͱ� �ƴ� �Ķ���͵��� ��ķ� �Ѿ�� ������ 0��° ��Ҹ��� �����ϰ� �Ѵ�.
			if (m_arrayInputVals[i]->isRange() || m_arrayInputVals[i]->isArray())
				getValInFunc(vals[i], m_arrayInputVals[i], 0);
			else
				vals[i]->copy(m_arrayInputVals[i]);
		}
	}

	// ������ ����ϱ�
	token_backup = token->evaluate(m_evaluator);

	// ������� �����ϱ�
	if (m_evaluator->m_val->m_array->getValue(0, 0)->isError()) {
		// �������� ������ �߻��Ͽ��ٸ� ����� 0��° ��ҿ� ��������
		// �����ǿ������Ƿ� �װ��� �������⿡ �����Ѵ�.
		QValueArray* srcRow = m_evaluator->m_val->m_array->getRow(0);
		for (int r = 0; r < nRow; r++) {
			QValueArray* dstRow = m_arrayResultVals->getRow(r);
			for (int c = 0; c < nCol; c++) {
				(*dstRow)[c]->BrCopy((*srcRow)[0]);
			}
		}
	}
	else {
		int nIndex = 0; // ��ȯ����
		int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
		while(nIndex < nTotalCount) {
			nRow = nIndex / m_arrayResultVals->getColCount();
			nCol = nIndex % m_arrayResultVals->getColCount();
			xlsValue* pValue = m_arrayResultVals->getValue(nRow, nCol);
			getResultValInFunc(pValue, m_evaluator->m_val->m_array, nIndex);

			nIndex++;
		}
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// ������ �������� ����� ��� �ش� ����� ����Լ�����ɿ��� �����ϱ� ���� �Լ�
// getValInFunc()�Լ��� ������.
void xlsArrayEvaluator::getResultValInFunc(xlsValue* pDst, xlsValueArray* pSrcArray, int nIndex)
{
	// ����Ŀ� ������ �� �� �Ĺ�ȣ�� ���
	int c_rows = m_arrayResultVals->getRowCount(); // cell�׷��� �ళ��(�Ƿ�, A1:D10���� v_rows = 10)
	int c_cols = m_arrayResultVals->getColCount(); // cell�׷��� �İ���(�Ƿ�, A1:D10���� v_cols = 4)
	int a_r = nIndex / c_cols;
	int a_c = nIndex % c_cols;

	int rows = pSrcArray->getRowCount();
	int cols = pSrcArray->getColCount();

	// Case 1 - cell�׷��� ũ�Ⱑ ����İ� ���ٸ�
	if (rows == c_rows && cols == c_cols) {
		xlsValue* v = pSrcArray->getValue(a_r, a_c);
		pDst->BrCopy(v);
	}
	else { // ������� ũ��� ���̳��� ��� spec�� [18.17.2.7] [Single- and Array Formulas]�� �����ȴ�� 
		// �Ķ���͵鿡 ���� ó���� �����Ѵ�.

		// ������� 1*1�����̶�� 1�� cell�� �����Ȱ�ó�� ����
		if (rows == 1 && cols == 1) {
			xlsValue* v = pSrcArray->getValue(0, 0);
			pDst->BrCopy(v);
		}
		// Case 2 - ���� cell�׷��� ���麸�� �� ���� ����� �����ٸ� ������ �� �������(left-most columns)�� cell�鿡 �����ȴ�.
		else if (c_rows < rows && c_cols >= cols && a_r >= c_rows) {
			xlsValue* v = pSrcArray->getValue(a_r, 0);
			pDst->BrCopy(v);
		}
		// Case 3 - ���� cell�׷��� ���麸�� �� ���� �ļ��� �����ٸ� ������ �ǿ��ʷĵ�(top-most rows)�� cell�鿡 �����ȴ�.
		else if (c_cols < cols && c_rows >= rows && a_c >= c_cols) {
			xlsValue* v = pSrcArray->getValue(0, a_c);
			pDst->BrCopy(v);
		}
		// Case 4 - ���� cell�׷��� ���麸�� ���� ����� �����ٸ� �� cell�� ������ ��츦 �����ϰ� �ڱ��� �����ġ�� �ش��� ���� ������.
		else if (c_rows >= rows && a_r >= rows) {
			// Case 4:1 - 1*N Ȥ�� 2�������� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� �ǿ�����cell���� �������� ���� ��(N/A)�� ������.
			if (a_c >= cols) {
				pDst->setError(eNA);
			}
			else if (c_rows >= 1 && rows > 1) {
				pDst->setError(eNA);
			}
			// Case 4:2 - N*1�� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� ����� ù��° ���� �����Ѵ�.
			else if (c_cols == 1) {
				xlsValue* v = pSrcArray->getValue(0, a_c);
				pDst->BrCopy(v);
			}
			else { // ��Ÿ
				if (rows == 1) {
					xlsValue* v = pSrcArray->getValue(0, a_c);
					pDst->BrCopy(v);
				}
				else {
					xlsValue* v = pSrcArray->getValue(a_r, a_c);
					pDst->BrCopy(v);
				}
			}
		}
		// Case 5 - ���� cell�׷��� ���麸�� ���� �ĵ��� �����ٸ� �� cell�� ������ ��츦 �����ϰ� �ڱ��� �����ġ�� �ش��� ���� ������.
		else if (c_cols >= cols && a_c >= cols) {
			// Case 5:1 - N*1 Ȥ�� 2�������� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� �ǹ��� cell���� �������� ���� ��(N/A)�� ������.
			if (a_r >= rows) {
				pDst->setError(eNA);
			}
			else if (c_cols >= 1 && cols > 1) {
				pDst->setError(eNA);
			}
			// Case 5:2 - 1*N�� cell�׷쿡 ���Ͽ� �ʰ��Ǵ� �ĵ��� ù��° ���� �����Ѵ�.
			else if (c_rows == 1) {
				xlsValue* v = pSrcArray->getValue(a_r, 0);
				pDst->BrCopy(v);
			}
			else { // ��Ÿ
				if (cols == 1) {
					xlsValue* v = pSrcArray->getValue(a_r, 0);
					pDst->BrCopy(v);
				}
				else {
					xlsValue* v = pSrcArray->getValue(a_r, a_c);
					pDst->BrCopy(v);
				}
			}
		}
		// Case 4�� 5�� ������ 
		else {
			xlsValue* v = pSrcArray->getValue(a_r, a_c);
			pDst->BrCopy(v);
		}
	}
}

// �Ķ���Ͱ� �����̰� ���� �Ѱ��� �Ķ���͸� ������ ������ �Ķ���͵��� Array(Ȥ�� Range)�� ���ĵ鿡 ���� ó����
// �� ó���ΰ� �����ϴ� ���ĵ� : SUBTOTAL
// nFixedIndex : ����� �ƴ� �Ķ������ ���ΰ�(0 based index)
xlsToken* xlsArrayEvaluator::doAnyOneNoneArrayFuncVar(xlsToken* token, BrINT nFixedIndex)
{
	// �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0; // ��ȯ����
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			// AddIn���� Ư���� ������ ��Ÿ���� ������ �Ķ���ͷĿ� ���Եǿ� ���´�.
			if (token->getFuncNum() == xlsFunc::eAddIn && m_arrayInputVals[i]->isFunc()) {
				vals[i]->copy(m_arrayInputVals[i]);
			}
			else if (i == nFixedIndex) {
				getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
			}
			else {
				vals[i]->copy(m_arrayInputVals[i]);
			}
		}

		token_backup = token->evaluate(m_evaluator);

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �Ķ������ �缳��
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// [��ļ��ı���-3�ܰ�] ��������� �Ķ���͵��� �������� �ʴ� ���ĵ��� ó��
// �Ϻ� ���ĵ��� spec�� ���� ����Ķ���͸� �������� ������ ��ļ�������ÿ���
// �ԷµǴ� �Ķ���͵��� ����Ķ���͵�� �ν������ν� �ᱹ ��ļ����� ��Ȯ�� ������� �ʴ´�.
// �׷��� ���ĵ��� �Ϻδ� ��ļ������������ ����Ķ������ ũ�⸦ �Ѿ�� ���
// �Ѿ�� ������ �κ��� #NA�� �����ϸ� �Ϻδ� ��ļ���������� ��ü�� #VALUE!�� �����Ѵ�.
// ��ļ������������ #VALUE!�� �����ϴ� ���ĵ鿡 ���ؼ��� �� �Լ��� ��ļ������������ #VALUE!��
// �����ϴ� ��ɵ� �Բ� �����Ѵ�.
// [�Ķ����]
// token : ó���� token
// eKind : xlsFunc::eFuncArgs::eAddIn���� ǥ���Ǵ� token���� �����ϴ� ������ ����
// nID : ó���� ������ ID
// [������]
// TRUE : ��������� �Ķ���͵��� �����ϴ� �����̴�.
// FALSE : ��������� �Ķ���͵��� �������� �ʴ� �����̴�.
//		  �ʿ� - ��������� �Ķ���͵��� �������� �ʴ� �����̶�� ������ ��ļ���������� ��ü�� #VALUE!��
//				�����ϴ� ������ ��� �� �Լ������� ó������ �����ϹǷ� FALSE�� �����ְ� �Ѵ�.
xlsArrayEvaluator::ENUM_SUPPORT_RESULT xlsArrayEvaluator::checkSupportArrayParameter(xlsToken* token, ENUM_ADDIN_KIND eKind, int nID)
{
	ENUM_SUPPORT_RESULT bRet = SUPPORTED;
	BIntArray fixedArrayParaIndexList; // spec�䱸�� ��ķ� �����Ǵ� �Ķ������ ��ȣ���� ���
	BrBOOL bArrayParameter = BrFALSE; // �Ϻ� ���ĵ��� �ɿ����� �����Ǵ� ����� �������� ������ �����ķμ� �����Ǵ� ����� �����Ѵ�.

	// ���ĺ��� �������θ� �˻��ϱ�
	if (eKind == ADDIN_BOND) {
		switch(nID) {
			// #VALUE!���� �����ϴ°�!
		case eYieldmat:
		case eYielddisc:
		case eYield:
		case eTbillyield:
		case eTbillprice:
		case eTbilleq:
		case eReceived:
		case ePricemat:
		case ePricedisc:
		case ePrice:
		case eOddlyield:
		case eOddlprice:
		case eOddfyield:
		case eOddfprice:
		case eIntrate:
		case eMduration:
		case eNominal:
		case eDuration:
		case eEffect:
		case eDisc:
		case eCumprinc:
		case eCumipmt:
		case eCouppcd:
		case eCoupnum:
		case eCoupncd:
		case eCoupdaysnc:
		case eCoupdays:
		case eCoupdaysbs:
		case eAmorlinc:
		case eAmordegrc:
		case eAccrintm:
		case eAccrint:
		case eYearFrac:
		case eEomonth:
		case eEdate:
			bRet = UNSUPPORT_UNPROCESS;
			break;
		default:
			bRet = SUPPORTED;
		}
	}
	else if (eKind == ADDIN_ENGINEERING) {
		switch(nID) {
		case eImtan:
		case eImsub:
		case eImsqrt:
		case eImsinh:
		case eImsin:
		case eImsec:
		case eImsech:
		case eImreal:
		case eImpower:
		case eImlog2:
		case eImlog10:
		case eImln:
		case eImexp:
		case eImdiv:
		case eImcsch:
		case eImcsc:
		case eImcot:
		case eImcosh:
		case eImcos:
		case eImconjugate:
		case eImargument:
		case eImaginary:
		case eImabs:
			bRet = UNSUPPORT_UNPROCESS;
			bArrayParameter = BrTRUE;
			break;
		case eErfc_Precise:
		case eErf_Precise:
		case eErfc:
		case eErf:
		case eComplex:
		case eBesselY:
		case eBesselK:
		case eBesselJ:
		case eBesselI:
			bRet = UNSUPPORT_UNPROCESS;
			break;
		default:
			bRet = SUPPORTED;
		}
	}
	else if (eKind == ADDIN_MISC) {
		switch(nID) {
			// #NA�� �����ϴ°�!
		case eNetWorkDays_Intl:
		case eWorkday_Intl:
			bRet = UNSUPPORT_UNPROCESS;
			return bRet;

			// #VALUE!���� �����ϴ°�!
		case eXirr:
			bRet = UNSUPPORT_UNPROCESS;
			fixedArrayParaIndexList.Add(1);
			fixedArrayParaIndexList.Add(2);
			break;
		case eXnpv:
			bRet = UNSUPPORT_UNPROCESS;
			fixedArrayParaIndexList.Add(2);
			fixedArrayParaIndexList.Add(3);
			break;
		case eFVSchedule:
			bRet = UNSUPPORT_UNPROCESS;
			fixedArrayParaIndexList.Add(2);
			break;
		case eConvert:
		case eDec2oct:
		case eDec2hex:
		case eDec2bin:
		case eBin2oct:
		case eBin2hex:
		case eBin2dec:
		case eOct2hex:
		case eOct2dec:
		case eOct2bin:
		case eHex2oct:
		case eHex2dec:
		case eHex2bin:
			bRet = UNSUPPORT_UNPROCESS;
			bArrayParameter = BrTRUE;
			break;
		case eSeriesSum:
			bRet = UNSUPPORT_UNPROCESS;
			fixedArrayParaIndexList.Add(4);
			break;
		case eRandbetween:
		case eQuotient:
		case eMRound:
		case eSqrtPI:
		case eFactDouble:
		case eIsOdd:
		case eIsEven:
		case eDollarfr:
		case eDollarde:
		case eGEStep:
		case eDelta:
		case eWeeknum:
			bRet = UNSUPPORT_UNPROCESS;
			break;
		case eNetWorkDays:
		case eWorkday:
			bRet = UNSUPPORT_UNPROCESS;
			fixedArrayParaIndexList.Add(3);
			break;
		default:
			bRet = SUPPORTED;
		}
	}
	else {
		xlsFunc::eFuncArgs eFuncNum = (xlsFunc::eFuncArgs)token->getFuncNum();
		switch(eFuncNum) {
		case xlsFunc::eIndirect:
			bRet = UNSUPPORT_UNPROCESS;
			break;
		default:
			bRet = SUPPORTED;
		}

		// xlsFunc::eFuncArgs::eAddIn���� ǥ������ �ʴ� ���ĵ鿡 ���ؼ��� �׳� ������� return�ϱ�
		return bRet;
	}

	if (bRet == SUPPORTED)
		return bRet;

	/********** �������� �ʴ� ���ĵ鿡 ���� ó�� ***********/
	// 1. �Լ��� �Ķ���������� ���
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// 2. �Ķ���͵��� ��� �ϳ��� ������� �Ǵ�.
	BrBOOL bArrayFound = BrFALSE;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		if (token->getFuncNum() == xlsFunc::eAddIn && pVal->isFunc())
			continue;

		int j = 0;
		for (j = 0; j < fixedArrayParaIndexList.size(); j++) {
			if (i == fixedArrayParaIndexList[j]) {
				j = -1;
				break;
			}
		}

		if (j == -1)
			continue;

		if (bArrayParameter && pVal->isArray()) {
			continue;
		}
		else if (checkArrayValue(pVal)) {
			bArrayFound = BrTRUE;
			break;
		}
	}

	// ����Ķ���͸� �ϳ��� �����ϰ� ���� �ʴٸ� ó�������Ѱ����� ���� return�ϱ�
	if (!bArrayFound)
		return SUPPORTED;

	// 3. ��ļ������������ #VALUE!�� �����ϱ�
	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nRow = m_arrayResultVals->getRowCount();
	int nCol = m_arrayResultVals->getColCount();
	xlsToken* token_backup = BrNULL;

	// ����� �����ϴ� ��ĺ����� NA�� �ʱ�ȭ�ϱ�
	for (int r = 0; r < nRow; r++) {
		for (int c = 0; c < nCol; c++) {
			m_arrayResultVals->getValue(r, c)->setError(eInvalidValue);
		}
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return UNSUPPORT_PROCESSED;
}

// �ش� ���� ����� �����ϰ��ִ��� �˻��ϴ� �Լ�
BrBOOL xlsArrayEvaluator::checkArrayValue(xlsCalValue* pVal)
{
	BrBOOL bArray = BrFALSE;
	if (pVal == BrNULL)
		return bArray;

	if (pVal->isArray()) {
		xlsValueArray* va = pVal->m_array;
		int rows = va->getRowCount();
		int cols = va->getColCount();

		if (rows > 1 || cols > 1)
			bArray = BrTRUE;
	}
	else if (pVal->isRange()) {
		xlsTRange rng;
		pVal->getRange(rng);
		int rows = rng.getNrRows();
		int cols = rng.getNrCols();

		if (rows > 1 || cols > 1)
			bArray = BrTRUE;
	}
	else {
		bArray = BrFALSE;
	}

	return bArray;
}

// CHOOSE���Ŀ� ���� ó����
// �⺻�����δ� doTokenFuncVar()�Լ��� ó�������� ����.
// ���� CHOOSE������ ��������� �Ϲݰ��ĵ�� ���̳��Ƿ� ���� ó���θ� �������.
xlsToken* xlsArrayEvaluator::doTokenChooseFuncVar(xlsToken* token)
{
	// �Լ��� �Ķ���������� ���
	xlsToken* token_backup = BrNULL;
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = 1; // ù��° �Ķ���͸� token���� �������ְ� �������� �̹� ��ĺ����� ����������.
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // ù��° �Ķ������ ���ΰ�

	// �Է��Ķ���͵��� �����ϱ� ���� ������� â��
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// �븮�������� �ʱ�ȭ
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// �Է��Ķ���͵��� Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// ���������� ũ��Ȯ��
	setResultBuffer();

	// ����Ķ���Ϳ� ���� ����ó��
	int nIndex = 0; // ��ȯ����
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // ��ȯ������ ũ��
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
		}

		token_backup = token->evaluate(m_evaluator);
		// ����� �Ķ���Ϳ� ���� ���
		if (token_backup) {
			token_backup = token_backup->evaluate(m_evaluator);
		}

		// ��ó�� : ���� ���� range�� array��� �װ����κ��� ���� ������� ǥ���� cell��ġ�� �����ϴ�
		// ����� ���� �򵵷� �Ѵ�.
		if (m_evaluator->m_val->isRange() || m_evaluator->m_val->isArray()) {
			xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
			getValInFunc(pCalValue, m_evaluator->m_val, nIndex);
			m_evaluator->m_val->copy(pCalValue);
			BR_SAFE_DELETE(pCalValue);
		}

		// ������� � cell�� ���� �������� ��� ǥ�þȵǴ� ������ ����.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// ������� �����ϱ�
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// �Ķ������ �缳��
		// SHEETS�� ���� �Ķ���Ͱ����� 0~1��(����)�� ��쵵 �����ϹǷ�...
		if (nArgCount > 0) {
			m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
		}
		else {
			m_evaluator->m_val = (*m_evaluator->m_vals)[0];
		}
	}

	// ��Ŀ������� ù��° �Ķ���͸� ���� ����⿡ �����ϱ�
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

#endif // USE_ARRAYFUNCTION_DANDONG
