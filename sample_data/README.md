# 공개용 가상 데이터

이 폴더의 파일은 GitHub 공개 저장소에서 바로 테스트할 수 있도록 만든 가상 근태 데이터입니다. 실제 인물, 조직, 사번, 근태 기록과 무관합니다.

## 파일

- `public_attendance_sample.xlsx`: 웹 화면에 업로드할 수 있는 샘플 엑셀
- `public_attendance_sample_expected.json`: 샘플 분석 기대 결과

## 재생성

```powershell
python scripts\create_public_sample.py
```

재생성 후 `python app.py`를 실행하고 `public_attendance_sample.xlsx`를 업로드하면 이상치와 검토 대상이 표시됩니다.
