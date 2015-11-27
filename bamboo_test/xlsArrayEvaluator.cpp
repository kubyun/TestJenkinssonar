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

	// [배렬수식처리-2] 배렬수식연산의 결과를 보관하기 위한 림시완충기의 초기화
	m_arrayResultVals = BrNULL;

	// [배렬수식처리-2] 배렬수식연산에 리용되는 파라메터들을 위한 림시변수들의 초기화
	m_arrayInputVal1 = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
	m_arrayInputVal2 = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
	m_arrayInputVal3 = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
	m_arrayInputVal4 = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
	m_arrayInputVal5 = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);

	// main cell의 초기화
	m_pCalcCell = BrNULL;
}

xlsArrayEvaluator::~xlsArrayEvaluator()
{
	// [배렬수식처리-2] 배렬수식연산에 리용된 완충기들의 해방
	BR_SAFE_DELETE(m_arrayResultVals);
	BrDELETE m_arrayInputVal1;
	BrDELETE m_arrayInputVal2;
	BrDELETE m_arrayInputVal3;
	BrDELETE m_arrayInputVal4;
	BrDELETE m_arrayInputVal5;

	// 배렬수식연산에 리용된 완충기들의 해방
	if (m_arrayInputVals.GetSize() > 0) {
		for (int i = 0; i < m_arrayInputVals.GetSize(); i++) {
			xlsCalValue* pValue = m_arrayInputVals[i];
			BR_SAFE_DELETE(pValue);
		}
	}
}

void xlsArrayEvaluator::recalcArrayFormula(xlsCalcCell* cell)
{
	// cell설정
	m_pCalcCell = cell;

	// 배렬수식적용령역을 판단하기 즉 배렬수식이 단일cell에 적용되는지 다중cell에 적용되는지 판단하기.
	// 다중cell에 적용된 경우 입력배렬과 적용범위의 크기가 같아야 한다.
	BrBOOL bSignleCell = BrTRUE; // 기정으로는 단일cell에 적용되는것으로 본다.
	int nCols = cell->m_arrayRef.getNrCols();
	int nRows = cell->m_arrayRef.getNrRows();
	if (nCols > 1 || nRows > 1)
		bSignleCell = BrFALSE;
	else
		bSignleCell = BrTRUE;

	// [배렬수식처리-2] 4칙연산자, 함수 등에 대해서는 입력파라메터와 배렬수식적용령역에 따르는 
	// 처리가 기존방식과는 다르게 순환처리되여야 한다.
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
			// 3단계(기본함수구현)에서 Sum기능구현과 관련하여 아래의 코드와 겹치므로
			// 아래의 코드를 주석처리한다.
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
			// 함수구현이 배렬함수를 적용할수 없게 되여있음. 따라서 이 함수에 대한 재구현이 필요함.
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
			// 4단계에서 처리되여야 할 내용.
			//token = doTokenFuncVar(token);
			//break;
		case m_eXLS_TokenGE:
		case m_eXLS_TokenEQ:
		case m_eXLS_TokenGT:
		case m_eXLS_TokenLE:
		case m_eXLS_TokenLT:
		case m_eXLS_TokenNE:
			// 4단계에서 처리되여야 할 내용.
			token = token->m_next;
			break;
		default:
			token = token->evaluate(m_evaluator);
			continue;
		}
	}

	BR_SAFE_DELETE(m_arrayResultVals);
}

// [배렬수식처리-2] 배렬수식에 대한 4칙연산처리
// 4칙연산에 대한 배렬처리이므로 입력파라메터의 개수가 2개라는것을 전제로 한다.
// token : 4칙연산에 대한 token
// 돌림값 : 4칙연산의 다음 token
xlsToken* xlsArrayEvaluator::processNumericalExpression(xlsToken* token)
{
	xlsCalValue *val1 = NULL, *val2 = NULL;
	xlsToken* token_backup = NULL;

	// 초기화
	val1 = m_evaluator->m_val->m_prev;
	val2 = m_evaluator->m_val;

	// 첫번째 파라메터의 Backup
	m_arrayInputVal1->copy(val1);

	// 두번째 파라메터의 Backup
	m_arrayInputVal2->copy(val2);

	// 연산조건이 성립하는가를 따져보기
	bool bCheck = checkArrayFromulaCondition(val1, val2);
	if (!bCheck) {
		(*m_evaluator->m_vals)[0]->setError(eNA);
		//		token = token->evaluate(this);
		return BrNULL;
	}

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nIndex < nTotalCount) {
		getValFromInputParameter(val1, m_arrayInputVal1, nIndex);
		getValFromInputParameter(val2, m_arrayInputVal2, nIndex);

		token_backup = token->evaluate(m_evaluator);

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 재설정
		m_evaluator->m_val = val2;
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기(val1)에 복사하기
	// 왜냐하면 배렬연산을 위한 파라메터가 n개인 경우(2개이상)가 있기때문이다.
	int nVal = val1->m_nVal;
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// [배렬수식처리-2] 배렬수식의 연산조건이 성립되는가를 검사하는 함수
// val1 : 첫번째 입력파라메터
// val2 : 두번째 입력파라메터
// 돌림값 : true-조건성립, false-조건부족
bool xlsArrayEvaluator::checkArrayFromulaCondition(xlsCalValue * val1, xlsCalValue * val2)
{
	// 두개의 파라메터가 배렬 혹은 령역인가를 검사하기
	bool bFlag1 = (val1->isRange() || val1->isArray());
	bool bFlag2 = (val2->isRange() || val2->isArray());

	// 첫번째 파라메터의 정보를 얻기
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

	// 두번째 파라메터의 정보를 얻기
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

	//두개의 파라메터중 어느 하나라도 배렬 혹은 령역이 아니라면
	if (!bFlag1 || !bFlag2) {
		// 두개의 파라메터중 하나가 배렬이라면 그 배렬의 크기를 알아내기
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

	// 두 파라메터에 대한 비교
	if (nRows1 == nRows2 && nCols1 == nCols2) {
		if (nRows1 > m_arrayResultVals->getRowCount() || nCols1 > m_arrayResultVals->getColCount())
			m_arrayResultVals->setSize(nRows1, nCols1);

		return true;
	}
	else {
		// 2개의 인수가 다 1차원배렬이라면
		if (nRows1 == 1 && nRows1 == nCols2) {
			m_arrayResultVals->setSize(nRows2, nCols1);
			return true;
		}

		if (nCols1 == 1 && nCols1 == nRows2) {
			m_arrayResultVals->setSize(nRows1, nCols2);
			return true;
		}

		// row개수가 같고 col개수는 1의 배수관계에 있다면
		int nMaxCol = BrMAX(nCols1, nCols2);
		int nMinCol = BrMIN(nCols1, nCols2);
		if (nRows1 == nRows2 && nMinCol == 1 && (nMaxCol % nMinCol) == 0) {
			m_arrayResultVals->setSize(nRows1, nMaxCol);
			return true;
		}

		// col개수가 같고 row개수는 1의 배수관계에 있다면
		int nMaxRow = BrMAX(nRows1, nRows2);
		int nMinRow = BrMIN(nRows1, nRows2);
		if (nCols1 == nCols2 && nMinRow == 1 && (nMaxRow % nMinRow) == 0) {
			m_arrayResultVals->setSize(nMaxRow, nCols1);
			return true;
		}

		return false;
	}
}

// [배렬수식처리-2] 배렬수식적용령역의 크기를 result완충기의 크기로 설정하기 
void xlsArrayEvaluator::setResultBuffer()
{
	int nRows = 0, nCols = 0;
	xlsCalcCell* pCell = m_pCalcCell;
	xlsTRange rng;
	nRows = pCell->m_arrayRef.getNrRows();
	nCols = pCell->m_arrayRef.getNrCols();

	m_arrayResultVals->setSize(nRows, nCols);
}

// [배렬수식처리-2] 기본함수의 배렬함수계산시 주어진 입력파라메터에서 nIndex에 해당한 값을 얻기
// pDst : 얻어진 값을 보관하는 변수
// pSrc : 입력파라메터
// nIndex : 입력파라메터의 색인
// 돌림값 : 해당 token의 다음 token
void xlsArrayEvaluator::getValInFunc(xlsCalValue* pDst, xlsCalValue* pSrc, int nIndex)
{
	// 값배렬에 따르는 행 및 렬번호를 얻기
	int c_rows = m_arrayResultVals->getRowCount(); // cell그룹의 행개수(실례, A1:D10에서 v_rows = 10)
	int c_cols = m_arrayResultVals->getColCount(); // cell그룹의 렬개수(실례, A1:D10에서 v_cols = 4)
	int a_r = nIndex / c_cols;
	int a_c = nIndex % c_cols;

	if (pSrc->isRange()) {
		xlsTRange rng;
		pSrc->getRange(rng);
		int rows = rng.getNrRows();
		int cols = rng.getNrCols();

		// Case 1 - cell그룹의 크기가 값배렬과 같다면
		if (rows == c_rows && cols == c_cols) {
			pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1() + a_c);
		}
		else { // 값배렬의 크기와 차이나는 경우 spec의 [18.17.2.7] [Single- and Array Formulas]에 규정된대로 
			// 파라메터들에 대한 처리를 진행한다.

			// 값배렬이 1*1형식이라면 1개 cell로 지정된것처럼 생각
			if (rows == 1 && cols == 1) {
				pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1(), pSrc->getCol1());
			}
			// Case 2 - 만일 cell그룹이 값들보다 더 적은 행들을 가진다면 값들의 맨 왼쪽행들(left-most columns)이 cell들에 보관된다.
			else if (c_rows < rows && c_cols >= cols && a_r >= c_rows) {
				pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1());
			}
			// Case 3 - 만일 cell그룹이 값들보다 더 적은 렬수를 가진다면 값들의 맨웃쪽렬들(top-most rows)이 cell들에 보관된다.
			else if (c_cols < cols && c_rows >= rows && a_c >= c_cols) {
				pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1(), pSrc->getCol1() + a_c);
			}
			// Case 4 - 만일 cell그룹이 값들보다 많은 행들을 가진다면 매 cell은 다음의 경우를 제외하고 자기의 상대위치에 해당한 값을 가진다.
			else if (c_rows >= rows && a_r >= rows) {
				// Case 4:1 - 1*N 혹은 2차원렬의 cell그룹에 대하여 초과되는 맨오른쪽cell들은 규정되지 않은 값(N/A)을 가진다.
				if (a_c >= cols) {
					pDst->setError(eNA);
				}
				else if (c_rows >= 1 && rows > 1) {
					pDst->setError(eNA);
				}
				// Case 4:2 - N*1의 cell그룹에 대하여 초과되는 행들은 첫번째 행을 복제한다.
				else if (c_cols == 1) {
					pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1(), pSrc->getCol1() + a_c);
				}
				else { // 기타
					if (rows == 1) {
						pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1(), pSrc->getCol1() + a_c);
					}
					else {
						pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1() + a_c);
					}
				}
			}
			// Case 5 - 만일 cell그룹이 값들보다 많은 렬들을 가진다면 매 cell은 다음의 경우를 제외하고 자기의 상대위치에 해당한 값을 가진다.
			else if (c_cols >= cols && a_c >= cols) {
				// Case 5:1 - N*1 혹은 2차원렬의 cell그룹에 대하여 초과되는 맨밑의 cell들은 규정되지 않은 값(N/A)을 가진다.
				if (a_r >= rows) {
					pDst->setError(eNA);
				}
				else if (c_cols >= 1 && cols > 1) {
					pDst->setError(eNA);
				}
				// Case 5:2 - 1*N의 cell그룹에 대하여 초과되는 렬들은 첫번째 렬을 복제한다.
				else if (c_rows == 1) {
					pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1());
				}
				else { // 기타
					if (cols == 1) {
						pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1());
					}
					else {
						pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1() + a_c);
					}
				}
			}
			// Case 4와 5의 정상경우 
			else {
				pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + a_r, pSrc->getCol1() + a_c);
			}
		}
	}
	else if (pSrc->isArray()) {
		xlsValueArray* va = pSrc->m_array;
		int rows = va->getRowCount();
		int cols = va->getColCount();

		// Case 1 - cell그룹의 크기가 값배렬과 같다면
		if (rows == c_rows && cols == c_cols) {
			xlsValue* v = va->getValue(a_r, a_c);
			pDst->BrCopy(v);
		}
		else { // 값배렬의 크기와 차이나는 경우 spec의 [18.17.2.7] [Single- and Array Formulas]에 규정된대로 
			// 파라메터들에 대한 처리를 진행한다.

			// 값배렬이 1*1형식이라면 1개 cell로 지정된것처럼 생각
			if (rows == 1 && cols == 1) {
				xlsValue* v = va->getValue(0, 0);
				pDst->BrCopy(v);
			}
			// Case 2 - 만일 cell그룹이 값들보다 더 적은 행들을 가진다면 값들의 맨 왼쪽행들(left-most columns)이 cell들에 보관된다.
			else if (c_rows < rows && c_cols >= cols && a_r >= c_rows) {
				xlsValue* v = va->getValue(a_r, 0);
				pDst->BrCopy(v);
			}
			// Case 3 - 만일 cell그룹이 값들보다 더 적은 렬수를 가진다면 값들의 맨웃쪽렬들(top-most rows)이 cell들에 보관된다.
			else if (c_cols < cols && c_rows >= rows && a_c >= c_cols) {
				xlsValue* v = va->getValue(0, a_c);
				pDst->BrCopy(v);
			}
			// Case 4 - 만일 cell그룹이 값들보다 많은 행들을 가진다면 매 cell은 다음의 경우를 제외하고 자기의 상대위치에 해당한 값을 가진다.
			else if (c_rows >= rows && a_r >= rows) {
				// Case 4:1 - 1*N 혹은 2차원렬의 cell그룹에 대하여 초과되는 맨오른쪽cell들은 규정되지 않은 값(N/A)을 가진다.
				if (a_c >= cols) {
					pDst->setError(eNA);
				}
				else if (c_rows >= 1 && rows > 1) {
					pDst->setError(eNA);
				}
				// Case 4:2 - N*1의 cell그룹에 대하여 초과되는 행들은 첫번째 행을 복제한다.
				else if (c_cols == 1) {
					xlsValue* v = va->getValue(0, a_c);
					pDst->BrCopy(v);
				}
				else { // 기타
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
			// Case 5 - 만일 cell그룹이 값들보다 많은 렬들을 가진다면 매 cell은 다음의 경우를 제외하고 자기의 상대위치에 해당한 값을 가진다.
			else if (c_cols >= cols && a_c >= cols) {
				// Case 5:1 - N*1 혹은 2차원렬의 cell그룹에 대하여 초과되는 맨밑의 cell들은 규정되지 않은 값(N/A)을 가진다.
				if (a_r >= rows) {
					pDst->setError(eNA);
				}
				else if (c_cols >= 1 && cols > 1) {
					pDst->setError(eNA);
				}
				// Case 5:2 - 1*N의 cell그룹에 대하여 초과되는 렬들은 첫번째 렬을 복제한다.
				else if (c_rows == 1) {
					xlsValue* v = va->getValue(a_r, 0);
					pDst->BrCopy(v);
				}
				else { // 기타
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
			// Case 4와 5의 정상경우 
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

// [배렬수식처리-2] 4칙연산의 배렬함수계산시 주어진 입력파라메터에서 nIndex에 해당한 값을 얻기
// pDst : 얻어진 값을 보관하는 변수
// pSrc : 입력파라메터
// nIndex : 입력파라메터의 색인
// 돌림값 : 4칙연산의 다음 token
void xlsArrayEvaluator::getValFromInputParameter(xlsCalValue* pDst, xlsCalValue* pSrc, int nIndex)
{
	//bool bContinue = false;
	// 값배렬에 따르는 행 및 렬번호를 얻기
	int v_rows = m_arrayResultVals->getRowCount();
	int v_cols = m_arrayResultVals->getColCount();
	int v_r = nIndex / v_cols;
	int v_c = nIndex % v_cols;

	if (pSrc->isRange()) {
		xlsTRange rng;
		pSrc->getRange(rng);
		int rows = rng.getNrRows();
		int cols = rng.getNrCols();

		// 배렬의 크기가 값배렬과 같다면
		if (rows == v_rows && cols == v_cols) {
			pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1() + v_r, pSrc->getCol1() + v_c);
		}
		else { // 값배렬의 크기와 차이난다면
			if (rows == 1) { // (1 * N)의 1차원행렬이라면
				pDst->setCell(m_evaluator->getSheet(), pSrc->getRow1(), pSrc->getCol1() + v_c);
			}
			else if (cols == 1) { // (N * 1)의 1차원행렬이라면
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

		// 배렬의 크기가 값배렬과 같다면
		if (rows == v_rows && cols == v_cols) {
			xlsValue* v = va->getValue(v_r, v_c);
			pDst->BrCopy(v);
		}
		else { // 값배렬의 크기와 차이난다면
			if (rows == 1) { // (1 * N)의 1차원행렬이라면
				xlsValue* v = va->getValue(0, v_c);
				pDst->BrCopy(v);
			}
			else if (cols == 1) { // (N * 1)의 1차원행렬이라면
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

// [배렬수식처리-2] 배렬수식에 대한 xlsTokenFuncBasic객체로 표현되는 함수처리(m_eXLS_TokenFuncBasic)
// xlsTokenFuncBasic함수에 대한 배렬처리이므로 입력파라메터의 개수가 1개 혹은 0개라는것을 전제로 한다.
// token : xlsTokenFuncBasic함수에 대한 token
// 돌림값 : xlsTokenFuncBasic함수의 다음 token
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

// 배렬수식에 대한 xlsTokenFuncBasic객체로 표현되는 함수처리(m_eXLS_TokenFuncBasic)
// 이 처리부가 지원하는 공식들 : (아직은 알수 없음)
// token : xlsTokenFuncBasic함수에 대한 token
// 돌림값 : xlsTokenFuncBasic함수의 다음 token
xlsToken* xlsArrayEvaluator::doTokenFuncBasic(xlsToken* token)
{
	xlsCalValue *val1 = NULL;
	xlsToken* token_backup = NULL;

	// 초기화
	val1 = m_evaluator->m_val;

	// 첫번째 파라메터의 Backup
	m_arrayInputVal1->copy(val1);

	// 결과를 보관하기 위한 변수조종
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

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nIndex < nTotalCount) {
		getValFromInputParameter(val1, m_arrayInputVal1, nIndex);

		token_backup = token->evaluate(m_evaluator);

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 재설정
		m_evaluator->m_val = val1;
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기(val1)에 복사하기
	// 왜냐하면 배렬연산을 위한 파라메터가 n개인 경우(2개이상)가 있기때문이다.
	int nVal = val1->m_nVal;
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// [배렬수식처리-2] 배렬수식에 대한 xlsTokenFunc객체로 표현되는 함수처리(m_eXLS_TokenFunc)
// xlsTokenFunc함수에 대한 배렬처리이므로 입력파라메터의 개수가 가변적이다.
// token : xlsTokenFunc함수에 대한 token
// nResultCount : xlsTokenFunc함수결과의 개수
// bSingCell : 단일cell인가를 나타내는 기발
// 돌림값 : xlsTokenFunc함수의 다음 token
xlsToken* xlsArrayEvaluator::processTokenFunc(xlsToken* token, int& nResultCount, BrBOOL bSingCell)
{
	xlsCalValue *val1 = NULL, *val2 = NULL;
	xlsToken* token_backup = NULL;

	// 초기화
	val1 = m_evaluator->m_val->m_prev;
	val2 = m_evaluator->m_val;
	nResultCount = 0;

	// 첫번째 파라메터의 Backup
	m_arrayInputVal1->copy(val1);

	// 두번째 파라메터의 Backup
	m_arrayInputVal2->copy(val2);

	// 결과를 보관하기 위한 변수조종
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

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nResultCount < nTotalCount) {
		nIndex = nResultCount;
		getValFromInputParameter(val2, m_arrayInputVal2, nIndex);

		token_backup = token->evaluate(m_evaluator);

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nResultCount / m_arrayResultVals->getColCount();
		nCol = nResultCount % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nResultCount++;

		// 재설정
		m_evaluator->m_val = val2;
		m_evaluator->m_val->m_prev->copy(m_arrayInputVal1);
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기(val1)에 복사하기
	// 왜냐하면 배렬연산을 위한 파라메터가 n개인 경우(2개이상)가 있기때문이다.
	int nVal = val1->m_nVal;
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 배렬수식에 대한 xlsTokenFunc객체로 표현되는 함수처리(m_eXLS_TokenFunc)
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

// 배렬수식에 대한 xlsTokenFuncVar객체로 표현되는 함수처리(m_eXLS_TokenFuncVar)
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

// eAddIn식별자를 가지는 공식들의 배렬수식 처리부(m_eXLS_TokenFuncVar)
xlsToken* xlsArrayEvaluator::processTokenAddInFunc(xlsToken* token)
{
	xlsToken* pNextToken = BrNULL;
	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	xlsCalValue* val = (*(m_evaluator->m_vals))[nVal];
	if (val->isFunc() == false) {
		return pNextToken;
	}

	// 공식이름을 얻기
	xlsFunc* pFunc = val->m_func;
	QString name(pFunc->m_name.data(), pFunc->m_name.size());

	// 공식이름을 비교하기
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
		// 빈 처리
	}

	return pNextToken;
}

// xlsEngineerFuncs클라스로 표현되는 공식들에 대한 배렬수식처리부
// 이 함수에서 처리하는 공식들의 실례 : IMPRODUCT, IMSUM 등
xlsToken* xlsArrayEvaluator::processTokenEngineeringFuncVar(xlsToken* token)
{
	xlsToken* pNextToken = BrNULL;

	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	xlsCalValue* val = (*(m_evaluator->m_vals))[nVal];
	if (val->isFunc() == false) {
		return pNextToken;
	}

	// [배렬수식구현-3단계] 배렬형식의 파라메터들을 지원하지 않는 공식들의 처리
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
	case eImproduct: // 함수구현에 오유가 있음.
	case eImsum:
		pNextToken = doTokenFuncVar(token, 1, BrTRUE);
		break;
	}

	return pNextToken;
}

// xlsEngineerFuncs클라스로 표현되는 공식들에 대한 배렬수식처리부
// 이 함수에서 처리하는 공식들의 실례 : EDATE, EOMONTH 등
xlsToken* xlsArrayEvaluator::processTokenBondFuncVar(xlsToken* token)
{
	xlsToken* pNextToken = BrNULL;

	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	xlsCalValue* val = (*(m_evaluator->m_vals))[nVal];
	if (val->isFunc() == false) {
		return pNextToken;
	}

	// [배렬수식구현-3단계] 배렬형식의 파라메터들을 지원하지 않는 공식들의 처리
	xlsBondFuncs* func = (xlsBondFuncs*)val->m_func;
	ENUM_SUPPORT_RESULT result = checkSupportArrayParameter(token, ADDIN_BOND, func->m_nID);
	if (result == UNSUPPORT_PROCESSED) {
		pNextToken = token->m_next;
		return pNextToken;
	}

	// 일반처리
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

// xlsMiscAddinFuncs클라스로 표현되는 공식들에 대한 배렬수식처리부
// 이 함수에서 처리하는 공식들의 실례 : ISOWEEKNUM, WORKDAY.INTL 등
xlsToken* xlsArrayEvaluator::processTokenMiscAddinFuncVar(xlsToken* token)
{
	xlsToken* pNextToken = BrNULL;

	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	xlsCalValue* val = (*(m_evaluator->m_vals))[nVal];
	if (val->isFunc() == false) {
		return pNextToken;
	}

	// [배렬수식구현-3단계] 배렬형식의 파라메터들을 지원하지 않는 공식들의 처리
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
/**************** Array Formula처리를 위한 기본공식들의 hanlder ************/
///////////////////////////////////////////////////////////////////////////

// 파라메터개수가 고정된 일반공식들에 대한 처리부(기정처리부)
// 대표적인 공식 : CONFIDENCE
// token : 해당 공식에 대한 token자료
// bIsAllArray : 모든 입력파라메터들이 spec에 따라 array인가를 나타내는 기발변수(기정값은 FALSE)
//				CHITEST, CHISQ.TEST등의 일부 공식들은 입력파라메터로 배렬을 요구한다.
xlsToken* xlsArrayEvaluator::doTokenFunc(xlsToken* token, BrBOOL bIsAllArray)
{
	// 함수의 파라메터정보를 얻기
	xlsTokenFunc* pToken = (xlsTokenFunc*)token;

	int nArgCount = (int)pToken->getFunc()->m_nMinArgs;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0; // 순환변수
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			if (bIsAllArray == BrTRUE)
				vals[i]->copy(m_arrayInputVals[i]);
			else
				getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
		}

		token_backup = token->evaluate(m_evaluator);
		// BITAND와 같은 일부 공식들에서는 결과값이 보관된 (*m_evaluator->m_vals)의 색인값(0)과 
		// m_evaluator->m_val의 색인값(1)이 일치하지 않는다.
		// 따라서 이에 대한 보정이 필요하다.(void xlsFunc::evaluate(xlsEvaluator* eval)함수를 참고할것)
		if (m_evaluator->m_val->m_nVal != nVal) {
			m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
		}

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 재설정
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 파라메터개수가 가변적인 일반공식들에 대한 처리부(기정처리부)
// 대표적인 공식 : BETADIST
// nFixParaCount : 고정되는 파라메터개수를 나타낸다. bIsAllRefArgs=TRUE일때에만 유효하다.
//				실례로 STDEV에 대해서는 nFixParaCount = 1, NPV에 대해서는 nFixParaCount = 2로 된다.
// bIsAllRefArgs : 입력파라메터들이 여러가지 자료형태(배렬, 참조, 값 등)의 값들이 설정될수 있는
//				1~255개까지의 가변적인 특성을 가지는가를 나타내는 기발변수(기정값은 FALSE)
//				TRUE로 설정하면 이와 같은 특성을 가지는 공식들(실례로 STDEV)을 처리한다.
xlsToken* xlsArrayEvaluator::doTokenFuncVar(xlsToken* token, BrINT nFixParaCount, BrBOOL bIsAllRefArgs)
{
	// [배렬수식구현-3단계] 배렬파라메터미지원 공식인가를 판단하기
	xlsToken* token_backup = BrNULL;
	BrBOOL bSupported = BrTRUE;
	ENUM_SUPPORT_RESULT result = checkSupportArrayParameter(token);
	if (result == UNSUPPORT_PROCESSED) { // 이미 처리되였다면
		token_backup = token->m_next;
		return token_backup;
	}
	else {
		if (result == SUPPORTED)
			bSupported = BrTRUE; // 지원하는 공식
		else
			bSupported = BrFALSE; // 지원하지 않는 공식
	}

	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0; // 순환변수
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
	while(nIndex < nTotalCount) {
		if (bIsAllRefArgs == BrTRUE) {
			for (int i = 0; i < nArgCount; i++) {
				// AddIn공식 특성상 공식을 나타내는 정보도 파라메터렬에 포함되여 들어온다.
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
				// AddIn공식 특성상 공식을 나타내는 정보도 파라메터렬에 포함되여 들어온다.
				if (token->getFuncNum() == xlsFunc::eAddIn && m_arrayInputVals[i]->isFunc()) {
					vals[i]->copy(m_arrayInputVals[i]);
				}
				else {
					getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
				}
			}
		}

		token_backup = token->evaluate(m_evaluator);

		// [배렬수식구현-3단계] 배렬형식의 파라메터들을 지원하지 않는 공식들의 처리
		// 여기서 처리되는 공식들은 배렬수식적용범위가 배렬파라메터의 크기를 넘어나는 경우
		// 넘어나는 나머지 부분이 #NA로 처리되게 해야 할 공식들에 대해서만 적용된다.
		if (!bSupported) {
			if (m_evaluator->m_val->isNA() == false)
				m_evaluator->m_val->setError(eInvalidValue);
		}
		else {
			// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
			if (m_evaluator->m_val->isCell()) {
				m_evaluator->m_val->checkValue(m_evaluator);
			}
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 파라메터의 재설정
		// SHEETS와 같이 파라메터개수가 0~1개(가변)인 경우도 존재하므로...
		if (nArgCount > 0) {
			m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
		}
		else {
			m_evaluator->m_val = (*m_evaluator->m_vals)[0];
		}
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 파라메터가 2개이고 첫번째 파라메터는 Array(혹은 Range)인 공식들에 대한 처리부
// 이 처리부가 지원하는 공식들 : SMALL, LARGE, PERCENTILE 등등
xlsToken* xlsArrayEvaluator::doFirstArray_Args2(xlsToken* token)
{
	xlsCalValue *val1 = NULL, *val2 = NULL;
	xlsToken* token_backup = NULL;

	// 초기화
	val1 = m_evaluator->m_val->m_prev;
	val2 = m_evaluator->m_val;

	// 첫번째 파라메터의 Backup
	m_arrayInputVal1->copy(val1);

	// 두번째 파라메터의 Backup
	m_arrayInputVal2->copy(val2);

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nIndex < nTotalCount) {
		getValInFunc(val2, m_arrayInputVal2, nIndex);

		token_backup = token->evaluate(m_evaluator);

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 재설정
		m_evaluator->m_val = val2;
		m_evaluator->m_val->m_prev->copy(m_arrayInputVal1);
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기(val1)에 복사하기
	// 왜냐하면 배렬연산을 위한 파라메터가 n개인 경우(2개이상)가 있기때문이다.
	int nVal = val1->m_nVal;
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 파라메터가 3개이고 첫번째 파라메터는 Array(혹은 Range), 3번째 파라메터는 optional인 공식들에 대한 처리부
// 이 처리부가 지원하는 공식들 : PERCENTRANK, PERCENTRANK.EXC, PERCENTRANK.INC
xlsToken* xlsArrayEvaluator::doFirstArray_Args3_Var(xlsToken* token)
{
	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	xlsCalValue *val1 = BrNULL, *val2 = BrNULL, *val3 = BrNULL;
	xlsToken* token_backup = BrNULL;

	// 초기화
	val1 = (*m_evaluator->m_vals)[nVal];
	val2 = (*m_evaluator->m_vals)[nVal + 1];

	// 첫번째 파라메터의 Backup
	m_arrayInputVal1->copy(val1);

	// 두번째 파라메터의 Backup
	m_arrayInputVal2->copy(val2);

	// 세번째 파라메터의 Backup
	if (nArgCount > 2) {
		val3 = (*m_evaluator->m_vals)[nVal + 2];
		m_arrayInputVal3->copy(val3);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nIndex < nTotalCount) {
		getValInFunc(val2, m_arrayInputVal2, nIndex);
		if (nArgCount > 2) {
			getValInFunc(val3, m_arrayInputVal3, nIndex);
		}

		token_backup = token->evaluate(m_evaluator);

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 재설정
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
		// 첫번째 파라메터를 그대로 복사
		(*m_evaluator->m_vals)[nVal]->copy(m_arrayInputVal1);
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기(val1)에 복사하기
	// 왜냐하면 배렬연산을 위한 파라메터가 n개인 경우(2개이상)가 있기때문이다.
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 파라메터가 n개로 고정이고 앞 혹은 뒤의 몇개 파라메터는 Array(혹은 Range)인 공식들에 대한 처리부
// 이 처리부가 지원하는 실례 공식들 : TTEST, T.TEST(4개고정, 첫번째, 두번째가 배렬)
// nArrayParameterCount : Array인 파라메터개수
// bForward : Array인 파라메터들이 파라메터렬의 앞쪽에 있는지 혹은 뒤쪽에 있는지 나타내는 기발변수
//			true : 앞쪽에 있음. false : 뒤에 있음.
xlsToken* xlsArrayEvaluator::doTokenFuncWithArray(xlsToken* token, BrINT nArrayParameterCount, BrBOOL bForward)
{
	// 함수의 파라메터정보를 얻기
	xlsTokenFunc* pToken = (xlsTokenFunc*)token;

	int nArgCount = (int)pToken->getFunc()->m_nMinArgs; 
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0; // 순환변수
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
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

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 재설정
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 파라메터가 3개로 고정이고 첫번째 및 세번째 파라메터들은 Array(혹은 Range)인 공식들에 대한 처리부
// 이 처리부가 지원하는 실례 공식들 : DAVERAGE, DCOUNT (대체로 자료기지관련 함수들)
xlsToken* xlsArrayEvaluator::doBothArray_Arg3(xlsToken* token)
{
	xlsCalValue *val1 = BrNULL, *val2 = BrNULL, *val3 = BrNULL;
	xlsToken* token_backup = BrNULL;

	// 함수의 파라메터정보를 얻기
	xlsTokenFunc* pToken = (xlsTokenFunc*)token;

	int nArgCount = (int)pToken->getFunc()->m_nMinArgs; 
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 초기화
	val1 = (*m_evaluator->m_vals)[nVal + 0];
	val2 = (*m_evaluator->m_vals)[nVal + 1];
	val3 = (*m_evaluator->m_vals)[nVal + 2];

	// 첫번째 파라메터의 Backup
	m_arrayInputVal1->copy(val1);

	// 두번째 파라메터의 Backup
	m_arrayInputVal2->copy(val2);

	// 세번째 파라메터의 Backup
	m_arrayInputVal3->copy(val3);

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0;
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount();
	while(nIndex < nTotalCount) {
		val1->copy(m_arrayInputVal1);
		getValInFunc(val2, m_arrayInputVal2, nIndex);
		val3->copy(m_arrayInputVal3);

		token_backup = token->evaluate(m_evaluator);

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 재설정
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기(val1)에 복사하기
	// 왜냐하면 배렬연산을 위한 파라메터가 n개인 경우(2개이상)가 있기때문이다.
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 파라메터가 2개로 고정이고 2개가 optional이며 2개의 optional 파라메터중
// 한개의 파라메터가 Array(혹은 Range)인 공식들에 대한 처리부
// 이 처리부가 지원하는 실례 공식들 : NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL
// bThirdArray : 세번째 파라메터 혹은 네번째 파라메터가 Array인가를 나타내는 기발변수
//				TRUE : 세번째 파라메터가 array, FALSE : 네번째 파라메터가 array
xlsToken* xlsArrayEvaluator::doWorkDaySerialFuncVar(xlsToken* token, BrBOOL bThirdArray)
{
	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값
	xlsCalValue* val = (*(m_evaluator->m_vals))[nVal];

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0; // 순환변수
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		BrBOOL bArrayFound = BrFALSE;

		for (int i = 0; i < nArgCount; i++) {
			// AddIn공식 특성상 공식을 나타내는 정보도 파라메터렬에 포함되여 들어온다.
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

			// bThirdArray변수를 리용하여 eNetWorkDays_Intl, eWorkday_Intl에 대해서만 
			// 모든 파라메터들이 배렬인가를 검사하기
			if (!bThirdArray && !bArrayFound && i != 4) {
				bArrayFound = checkArrayValue(m_arrayInputVals[i]);
			}
		}

		token_backup = token->evaluate(m_evaluator);

		// [배렬수식구현-3단계] 배렬형식의 파라메터들을 지원하지 않는 공식들의 처리
		// 여기서 처리되는 공식들은 배렬수식적용범위가 배렬파라메터의 크기를 넘어나는 경우
		// 넘어나는 나머지 부분이 #NA로 처리되게 해야 할 공식들에 대해서만 적용된다.
		if (!bThirdArray && bArrayFound) {
			if (m_evaluator->m_val->isNA() == false)
				m_evaluator->m_val->setError(eInvalidValue);
		}
		else {
			// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
			if (m_evaluator->m_val->isCell()) {
				m_evaluator->m_val->checkValue(m_evaluator);
			}
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 파라메터의 재설정
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 파라메터가 0개인 공식들에 대한 처리부
// 이 처리부가 지원하는 실례 공식들 : NOW, TODAY, RAND
xlsToken* xlsArrayEvaluator::doArgs0(xlsToken* token)
{
	int nVal = 0;

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0; // 순환변수
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		token_backup = token->evaluate(m_evaluator);

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 파라메터가 가변이고 그중 한개의 파라메터는 Array(혹은 Range)인 공식들에 대한 처리부
// 이 처리부가 지원하는 공식들 : FVSCHEDULE, RANK, RANK.AVG, RANK.EQ
// nArrayIndex : 배렬인 파라메터의 색인값(0 based index)
xlsToken* xlsArrayEvaluator::doAnyOneArrayFuncVar(xlsToken* token, BrINT nArrayIndex)
{
	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0; // 순환변수
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			// AddIn공식 특성상 공식을 나타내는 정보도 파라메터렬에 포함되여 들어온다.
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

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 파라메터의 재설정
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 파라메터가 가변이고 그중 앞의 몇개 파라메터는 Array(혹은 Range)인 공식들에 대한 처리부
// 이 처리부가 지원하는 공식들 : XIRR
// nLastArrayParaIndex : 배렬인 파라메터들중 마지막 파라메터의 색인(0 based index)
xlsToken* xlsArrayEvaluator::doSomeArraysFuncVar(xlsToken* token, BrINT nLastArrayParaIndex)
{
	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0; // 순환변수
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			// AddIn공식 특성상 공식을 나타내는 정보도 파라메터렬에 포함되여 들어온다.
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

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 파라메터의 재설정
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// COLMUN, ROW공식들에 대한 처리부
// 이 공식들은 파라메터가 0인 경우 배렬공식들이 적용되는 해당 cell들의 위치에 의존되므로
// 전용함수를 통한 례외처리가 필요하다.
xlsToken* xlsArrayEvaluator::doRowColFuncVar(xlsToken* token, BrBOOL bIsColumn)
{
	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0; // 순환변수
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();

		for (int i = 0; i < nArgCount; i++) {
			int m = 0;

			// COLUMN인 경우 입력파라메터의 크기에 상관없이 column개수에만 의존한다.
			if (bIsColumn) {
				m = nCol;
			}
			else { // ROW인 경우 입력파라메터의 크기에 상관없이 column개수에만 의존한다.
				m = nRow * m_arrayResultVals->getColCount();

			}
			getValInFunc(vals[i], m_arrayInputVals[i], m);
		}

		token_backup = token->evaluate(m_evaluator);

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		if (nArgCount > 0) {
			m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);
		}
		else {
			int nValue = m_evaluator->m_val->getNumber();
			if (bIsColumn) { // COLUMN공식이라면
				nValue = nValue + nCol;
			}
			else { // ROW공식이라면
				nValue = nValue + nRow;
			}

			m_arrayResultVals->getValue(nRow, nCol)->setValue(nValue);
		}

		nIndex++;

		// 파라메터의 재설정
		if (nArgCount > 0) {
			m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
		}
		else {
			m_evaluator->m_val = (*m_evaluator->m_vals)[0];
		}
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// HYPERLINK공식에 대한 처리부
// 두번째 파라메터가 Range로 되는 경우 Range의 (0, 0)에 해당한 값만을 리용한다.(견본화일을 볼것)
// 따라서 이에 대하여 전용함수를 통한 례외처리가 필요하다.
xlsToken* xlsArrayEvaluator::doHyperlinkFuncVar(xlsToken* token)
{
	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0; // 순환변수
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			if (i < 1)
				getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
			else
				getValInFunc(vals[i], m_arrayInputVals[i], 0);
		}

		token_backup = token->evaluate(m_evaluator);

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 파라메터의 재설정
		// SHEETS와 같이 파라메터개수가 0~1개(가변)인 경우도 존재하므로...
		if (nArgCount > 0) {
			m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
		}
		else {
			m_evaluator->m_val = (*m_evaluator->m_vals)[0];
		}
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 파라메터가 가변이고 그중 앞의 몇개 파라메터는 Array(혹은 Range)이고 돌림값이 배렬인 공식들에 대한 처리부
// 이 처리부가 지원하는 공식들 : TREND, LINEST, LOGEST, GROWTH
// nLastArrayParaIndex : 배렬인 파라메터들중 마지막 파라메터의 색인(0 based index)
xlsToken* xlsArrayEvaluator::doSomeArraysFuncVarWithReturnArray(xlsToken* token, BrINT nLastArrayParaIndex)
{
	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nRow = m_arrayResultVals->getRowCount();
	int nCol = m_arrayResultVals->getColCount();
	xlsToken* token_backup = BrNULL;

	// 결과를 보존하는 배렬변수를 NA로 초기화하기
	for (int r = 0; r < nRow; r++) {
		for (int c = 0; c < nCol; c++) {
			m_arrayResultVals->getValue(r, c)->setError(eNA);
		}
	}

	// 파라메터들을 설정하기
	for (int i = 0; i < nArgCount; i++) {
		// AddIn공식 특성상 공식을 나타내는 정보도 파라메터렬에 포함되여 들어온다.
		if (token->getFuncNum() == xlsFunc::eAddIn && m_arrayInputVals[i]->isFunc()) {
			vals[i]->copy(m_arrayInputVals[i]);
		}
		else if (i <= nLastArrayParaIndex) {
			vals[i]->copy(m_arrayInputVals[i]);
		}
		else {
			// 배렬파라메터가 아닌 파라메터들이 배렬로 넘어올 때에는 0번째 요소만을 선택하게 한다.
			if (m_arrayInputVals[i]->isRange() || m_arrayInputVals[i]->isArray())
				getValInFunc(vals[i], m_arrayInputVals[i], 0);
			else
				vals[i]->copy(m_arrayInputVals[i]);
		}
	}

	// 공식을 계산하기
	token_backup = token->evaluate(m_evaluator);

	// 계산결과를 보관하기
	if (m_evaluator->m_val->m_array->getValue(0, 0)->isError()) {
		// 계산과정에 오유가 발생하였다면 배렬의 0번째 요소에 오유값이
		// 보관되여있으므로 그것을 결과완충기에 복사한다.
		QValueArray* srcRow = m_evaluator->m_val->m_array->getRow(0);
		for (int r = 0; r < nRow; r++) {
			QValueArray* dstRow = m_arrayResultVals->getRow(r);
			for (int c = 0; c < nCol; c++) {
				(*dstRow)[c]->BrCopy((*srcRow)[0]);
			}
		}
	}
	else {
		int nIndex = 0; // 순환변수
		int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
		while(nIndex < nTotalCount) {
			nRow = nIndex / m_arrayResultVals->getColCount();
			nCol = nIndex % m_arrayResultVals->getColCount();
			xlsValue* pValue = m_arrayResultVals->getValue(nRow, nCol);
			getResultValInFunc(pValue, m_evaluator->m_val->m_array, nIndex);

			nIndex++;
		}
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 파라메터가 고정이고 그중 앞의 몇개 파라메터는 Array(혹은 Range)이고 돌림값이 배렬인 공식들에 대한 처리부
// 이 처리부가 지원하는 공식들 : FREQUENCY
// nLastArrayParaIndex : 배렬인 파라메터들중 마지막 파라메터의 색인(0 based index)
xlsToken* xlsArrayEvaluator::doSomeArraysFuncWithReturnArray(xlsToken* token, BrINT nLastArrayParaIndex)
{
	// 함수의 파라메터정보를 얻기
	xlsTokenFunc* pToken = (xlsTokenFunc*)token;
	int nArgCount = (int)pToken->getFunc()->m_nMinArgs;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nRow = m_arrayResultVals->getRowCount();
	int nCol = m_arrayResultVals->getColCount();
	xlsToken* token_backup = BrNULL;

	// 결과를 보존하는 배렬변수를 NA로 초기화하기
	for (int r = 0; r < nRow; r++) {
		for (int c = 0; c < nCol; c++) {
			m_arrayResultVals->getValue(r, c)->setError(eNA);
		}
	}

	// 파라메터들을 설정하기
	for (int i = 0; i < nArgCount; i++) {
		// AddIn공식 특성상 공식을 나타내는 정보도 파라메터렬에 포함되여 들어온다.
		if (token->getFuncNum() == xlsFunc::eAddIn && m_arrayInputVals[i]->isFunc()) {
			vals[i]->copy(m_arrayInputVals[i]);
		}
		else if (i <= nLastArrayParaIndex) {
			vals[i]->copy(m_arrayInputVals[i]);
		}
		else {
			// 배렬파라메터가 아닌 파라메터들이 배렬로 넘어올 때에는 0번째 요소만을 선택하게 한다.
			if (m_arrayInputVals[i]->isRange() || m_arrayInputVals[i]->isArray())
				getValInFunc(vals[i], m_arrayInputVals[i], 0);
			else
				vals[i]->copy(m_arrayInputVals[i]);
		}
	}

	// 공식을 계산하기
	token_backup = token->evaluate(m_evaluator);

	// 계산결과를 보관하기
	if (m_evaluator->m_val->m_array->getValue(0, 0)->isError()) {
		// 계산과정에 오유가 발생하였다면 배렬의 0번째 요소에 오유값이
		// 보관되여있으므로 그것을 결과완충기에 복사한다.
		QValueArray* srcRow = m_evaluator->m_val->m_array->getRow(0);
		for (int r = 0; r < nRow; r++) {
			QValueArray* dstRow = m_arrayResultVals->getRow(r);
			for (int c = 0; c < nCol; c++) {
				(*dstRow)[c]->BrCopy((*srcRow)[0]);
			}
		}
	}
	else {
		int nIndex = 0; // 순환변수
		int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
		while(nIndex < nTotalCount) {
			nRow = nIndex / m_arrayResultVals->getColCount();
			nCol = nIndex % m_arrayResultVals->getColCount();
			xlsValue* pValue = m_arrayResultVals->getValue(nRow, nCol);
			getResultValInFunc(pValue, m_evaluator->m_val->m_array, nIndex);

			nIndex++;
		}
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// 공식의 돌림값이 배렬인 경우 해당 결과를 배렬함수적용령역에 복사하기 위한 함수
// getValInFunc()함수를 참조함.
void xlsArrayEvaluator::getResultValInFunc(xlsValue* pDst, xlsValueArray* pSrcArray, int nIndex)
{
	// 값배렬에 따르는 행 및 렬번호를 얻기
	int c_rows = m_arrayResultVals->getRowCount(); // cell그룹의 행개수(실례, A1:D10에서 v_rows = 10)
	int c_cols = m_arrayResultVals->getColCount(); // cell그룹의 렬개수(실례, A1:D10에서 v_cols = 4)
	int a_r = nIndex / c_cols;
	int a_c = nIndex % c_cols;

	int rows = pSrcArray->getRowCount();
	int cols = pSrcArray->getColCount();

	// Case 1 - cell그룹의 크기가 값배렬과 같다면
	if (rows == c_rows && cols == c_cols) {
		xlsValue* v = pSrcArray->getValue(a_r, a_c);
		pDst->BrCopy(v);
	}
	else { // 값배렬의 크기와 차이나는 경우 spec의 [18.17.2.7] [Single- and Array Formulas]에 규정된대로 
		// 파라메터들에 대한 처리를 진행한다.

		// 값배렬이 1*1형식이라면 1개 cell로 지정된것처럼 생각
		if (rows == 1 && cols == 1) {
			xlsValue* v = pSrcArray->getValue(0, 0);
			pDst->BrCopy(v);
		}
		// Case 2 - 만일 cell그룹이 값들보다 더 적은 행들을 가진다면 값들의 맨 왼쪽행들(left-most columns)이 cell들에 보관된다.
		else if (c_rows < rows && c_cols >= cols && a_r >= c_rows) {
			xlsValue* v = pSrcArray->getValue(a_r, 0);
			pDst->BrCopy(v);
		}
		// Case 3 - 만일 cell그룹이 값들보다 더 적은 렬수를 가진다면 값들의 맨웃쪽렬들(top-most rows)이 cell들에 보관된다.
		else if (c_cols < cols && c_rows >= rows && a_c >= c_cols) {
			xlsValue* v = pSrcArray->getValue(0, a_c);
			pDst->BrCopy(v);
		}
		// Case 4 - 만일 cell그룹이 값들보다 많은 행들을 가진다면 매 cell은 다음의 경우를 제외하고 자기의 상대위치에 해당한 값을 가진다.
		else if (c_rows >= rows && a_r >= rows) {
			// Case 4:1 - 1*N 혹은 2차원렬의 cell그룹에 대하여 초과되는 맨오른쪽cell들은 규정되지 않은 값(N/A)을 가진다.
			if (a_c >= cols) {
				pDst->setError(eNA);
			}
			else if (c_rows >= 1 && rows > 1) {
				pDst->setError(eNA);
			}
			// Case 4:2 - N*1의 cell그룹에 대하여 초과되는 행들은 첫번째 행을 복제한다.
			else if (c_cols == 1) {
				xlsValue* v = pSrcArray->getValue(0, a_c);
				pDst->BrCopy(v);
			}
			else { // 기타
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
		// Case 5 - 만일 cell그룹이 값들보다 많은 렬들을 가진다면 매 cell은 다음의 경우를 제외하고 자기의 상대위치에 해당한 값을 가진다.
		else if (c_cols >= cols && a_c >= cols) {
			// Case 5:1 - N*1 혹은 2차원렬의 cell그룹에 대하여 초과되는 맨밑의 cell들은 규정되지 않은 값(N/A)을 가진다.
			if (a_r >= rows) {
				pDst->setError(eNA);
			}
			else if (c_cols >= 1 && cols > 1) {
				pDst->setError(eNA);
			}
			// Case 5:2 - 1*N의 cell그룹에 대하여 초과되는 렬들은 첫번째 렬을 복제한다.
			else if (c_rows == 1) {
				xlsValue* v = pSrcArray->getValue(a_r, 0);
				pDst->BrCopy(v);
			}
			else { // 기타
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
		// Case 4와 5의 정상경우 
		else {
			xlsValue* v = pSrcArray->getValue(a_r, a_c);
			pDst->BrCopy(v);
		}
	}
}

// 파라메터가 가변이고 그중 한개의 파라메터를 제외한 나머지 파라메터들이 Array(혹은 Range)인 공식들에 대한 처리부
// 이 처리부가 지원하는 공식들 : SUBTOTAL
// nFixedIndex : 배렬이 아닌 파라메터의 색인값(0 based index)
xlsToken* xlsArrayEvaluator::doAnyOneNoneArrayFuncVar(xlsToken* token, BrINT nFixedIndex)
{
	// 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0; // 순환변수
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
	xlsToken* token_backup = BrNULL;
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			// AddIn공식 특성상 공식을 나타내는 정보도 파라메터렬에 포함되여 들어온다.
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

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 파라메터의 재설정
		m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

// [배렬수식구현-3단계] 배렬형식의 파라메터들을 지원하지 않는 공식들의 처리
// 일부 공식들은 spec에 따라 배렬파라메터를 지원하지 않으며 배렬수식적용시에도
// 입력되는 파라메터들을 배렬파라메터들로 인식함으로써 결국 배렬수식이 정확히 적용되지 않는다.
// 그러한 공식들중 일부는 배렬수식적용범위가 배렬파라메터의 크기를 넘어나는 경우
// 넘어나는 나머지 부분은 #NA로 설정하며 일부는 배렬수식적용범위 전체를 #VALUE!로 설정한다.
// 배렬수식적용범위를 #VALUE!로 설정하는 공식들에 대해서는 이 함수는 배렬수식적용범위를 #VALUE!로
// 설정하는 기능도 함께 수행한다.
// [파라메터]
// token : 처리할 token
// eKind : xlsFunc::eFuncArgs::eAddIn으로 표현되는 token들이 포함하는 공식의 종류
// nID : 처리할 공식의 ID
// [돌림값]
// TRUE : 배렬형식의 파라메터들을 지원하는 공식이다.
// FALSE : 배렬형식의 파라메터들을 지원하지 않는 공식이다.
//		  례외 - 배렬형식의 파라메터들을 지원하지 않는 공식이라고 할지라도 배렬수식적용범위 전체를 #VALUE!로
//				설정하는 공식인 경우 이 함수내에서 처리까지 진행하므로 FALSE를 돌려주게 한다.
xlsArrayEvaluator::ENUM_SUPPORT_RESULT xlsArrayEvaluator::checkSupportArrayParameter(xlsToken* token, ENUM_ADDIN_KIND eKind, int nID)
{
	ENUM_SUPPORT_RESULT bRet = SUPPORTED;
	BIntArray fixedArrayParaIndexList; // spec요구상 배렬로 지정되는 파라메터의 번호들의 목록
	BrBOOL bArrayParameter = BrFALSE; // 일부 공식들은 령역으로 지정되는 배렬은 지원하지 않지만 상수배렬로서 지원되는 배렬은 지원한다.

	// 공식별로 지원여부를 검사하기
	if (eKind == ADDIN_BOND) {
		switch(nID) {
			// #VALUE!만을 포함하는것!
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
			// #NA를 포함하는것!
		case eNetWorkDays_Intl:
		case eWorkday_Intl:
			bRet = UNSUPPORT_UNPROCESS;
			return bRet;

			// #VALUE!만을 포함하는것!
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

		// xlsFunc::eFuncArgs::eAddIn으로 표현되지 않는 공식들에 대해서는 그냥 결과값만 return하기
		return bRet;
	}

	if (bRet == SUPPORTED)
		return bRet;

	/********** 지원하지 않는 공식들에 대한 처리 ***********/
	// 1. 함수의 파라메터정보를 얻기
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = (int)pToken->m_bArgCount;
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 2. 파라메터들중 어느 하나라도 배렬인지 판단.
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

	// 배렬파라메터를 하나도 포함하고 있지 않다면 처리가능한것으로 보고 return하기
	if (!bArrayFound)
		return SUPPORTED;

	// 3. 배렬수식적용범위를 #VALUE!로 설정하기
	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nRow = m_arrayResultVals->getRowCount();
	int nCol = m_arrayResultVals->getColCount();
	xlsToken* token_backup = BrNULL;

	// 결과를 보존하는 배렬변수를 NA로 초기화하기
	for (int r = 0; r < nRow; r++) {
		for (int c = 0; c < nCol; c++) {
			m_arrayResultVals->getValue(r, c)->setError(eInvalidValue);
		}
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return UNSUPPORT_PROCESSED;
}

// 해당 값이 배렬을 포함하고있는지 검사하는 함수
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

// CHOOSE공식에 대한 처리부
// 기본적으로는 doTokenFuncVar()함수와 처리내용이 같음.
// 단지 CHOOSE공식의 구현방법이 일반공식들과 차이나므로 따로 처리부를 만들었음.
xlsToken* xlsArrayEvaluator::doTokenChooseFuncVar(xlsToken* token)
{
	// 함수의 파라메터정보를 얻기
	xlsToken* token_backup = BrNULL;
	xlsTokenFuncVar* pToken = (xlsTokenFuncVar*)token;
	int nArgCount = 1; // 첫번째 파라메터만 token으로 가지고있고 나머지는 이미 배렬변수로 가지고있음.
	int nVal = m_evaluator->m_val->m_nVal - (nArgCount - 1); // 첫번째 파라메터의 색인값

	// 입력파라메터들을 보관하기 위한 완충기의 창조
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
		m_arrayInputVals.Add(pCalValue);
	}

	// 대리변수들의 초기화
	CalValueArray vals;
	for (int i = 0; i < nArgCount; i++) {
		xlsCalValue* pVal = (*m_evaluator->m_vals)[nVal + i];
		vals.Add(pVal);
	}

	// 입력파라메터들의 Backup
	for (int i = 0; i < nArgCount; i++) {
		m_arrayInputVals[i]->copy(vals[i]);
	}

	// 결과완충기의 크기확정
	setResultBuffer();

	// 배렬파라메터에 대한 연산처리
	int nIndex = 0; // 순환변수
	int nRow = 0, nCol = 0;
	int nTotalCount = m_arrayResultVals->getRowCount() * m_arrayResultVals->getColCount(); // 순환변수의 크기
	while(nIndex < nTotalCount) {
		for (int i = 0; i < nArgCount; i++) {
			getValInFunc(vals[i], m_arrayInputVals[i], nIndex);
		}

		token_backup = token->evaluate(m_evaluator);
		// 얻어진 파라메터에 대한 계산
		if (token_backup) {
			token_backup = token_backup->evaluate(m_evaluator);
		}

		// 값처리 : 얻은 값이 range나 array라면 그값으로부터 현재 결과값을 표현할 cell위치에 대응하는
		// 요소의 값을 얻도록 한다.
		if (m_evaluator->m_val->isRange() || m_evaluator->m_val->isArray()) {
			xlsCalValue* pCalValue = BrNEW xlsCalValue(m_evaluator, m_calcEngine->m_group);
			getValInFunc(pCalValue, m_evaluator->m_val, nIndex);
			m_evaluator->m_val->copy(pCalValue);
			BR_SAFE_DELETE(pCalValue);
		}

		// 결과값이 어떤 cell에 대한 참조값인 경우 표시안되는 오유의 수정.
		if (m_evaluator->m_val->isCell()) {
			m_evaluator->m_val->checkValue(m_evaluator);
		}

		// 계산결과를 보관하기
		nRow = nIndex / m_arrayResultVals->getColCount();
		nCol = nIndex % m_arrayResultVals->getColCount();
		m_arrayResultVals->getValue(nRow, nCol)->BrCopy(m_evaluator->m_val);

		nIndex++;

		// 파라메터의 재설정
		// SHEETS와 같이 파라메터개수가 0~1개(가변)인 경우도 존재하므로...
		if (nArgCount > 0) {
			m_evaluator->m_val = (*m_evaluator->m_vals)[nArgCount-1];
		}
		else {
			m_evaluator->m_val = (*m_evaluator->m_vals)[0];
		}
	}

	// 배렬연산결과를 첫번째 파라메터를 위한 완충기에 복사하기
	m_evaluator->m_val = (*m_evaluator->m_vals)[nVal];
	m_evaluator->m_val->makeArray()->copy(m_arrayResultVals);

	return token_backup;
}

#endif // USE_ARRAYFUNCTION_DANDONG
