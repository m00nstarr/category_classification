
#################################################################################
# coded by Mun Hyung Lee / Email : moonstar114@naver.com , Phone : 010-4001-9614

-------------------
code update v1.0.1 ( 20/07/14 )
1. output file name에 생성일시 추가
2. 프로그램 진행 과정 설명 코드 추가

-------------------

설명서 
1. 확장자를 제외한 파일 이름을 입력받습니다. 

2. inputfile은 보존되고 industry가 새로 새겨진RAW 데이터가 새로 생성됩니다.
new file name = (기존 파일 이름)_out_(연월일시초).xlsx


3. inputfile은 다음과 같은 조건을 따라야합니다.

1) 코드와 같은 폴더 내에 있어야 함
2) I열에 industry라는 이름이 있어야 합니다. (웬만하면 비어져 있는 것이 좋음)
3) 없는 파일인 경우에 에러를 호출하며 프로그램이 종료됩니다.


industry 로 rename 
input 몇개의 레코드 >> output 몇개다 echo 해주어야 함.
파일 개수로 확인을 다시 해주어야함!!
토탈 count 수를 처음부터 세고 / 환경이 고려 되어야 한다 >>  파일을 잘라서 실행 

4. 실행은 shell 에서
> python3.5 check_category.py 를 통해 실행가능합니다.

* pandas, numpy, openpyxl, xlrd 라이브러리를 사용하므로 반드시 pip install을 통해 패키지를 다운로드

5. 프로그램의 흐름
1) 원본 데이터를 outputfile로 복사
2) 복사를 하면서 null data 가 있는 어플리케이션과 sports의 어플리케이션을 기록
3) 복사가 완료 후에 sports의 게임과 스포츠가 혼합되어있는 것을 질의를 통해 분류
4) null data의 category를 직접 기입
5) 엑셀 파일 최종 저장

