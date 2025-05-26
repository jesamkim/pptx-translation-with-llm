import boto3
import json
from pptx import Presentation
from pptx.shapes.base import BaseShape
from pptx.shapes.autoshape import Shape
from pptx.shapes.placeholder import PlaceholderPicture
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.shapes.picture import Picture
from pptx.shapes.graphfrm import GraphicFrame
import os
from typing import List, Dict, Any

class PowerPointTranslator:
    def __init__(self):
        # Amazon Bedrock 클라이언트 설정
        self.bedrock_client = boto3.client(
            'bedrock-runtime',
            region_name='us-west-2'
        )
        self.model_id = "us.anthropic.claude-3-5-sonnet-20240620-v1:0"
        
        # 지원하는 언어 목록
        self.supported_languages = {
            'ko': '한국어',
            'en': '영어',
            'ja': '일본어',
            'zh': '중국어',
            'fr': '프랑스어',
            'de': '독일어',
            'es': '스페인어',
            'it': '이탈리아어',
            'pt': '포르투갈어',
            'ru': '러시아어'
        }
    
    def extract_text_from_slide(self, slide) -> List[Dict[str, Any]]:
        """슬라이드에서 텍스트 정보를 추출합니다."""
        text_items = []
        
        # 모든 도형 처리 (간단하게 최대한 많은 텍스트 추출)
        for shape_idx, shape in enumerate(slide.shapes):
            try:
                # 1. 일반 텍스트가 있는 경우 (가장 단순한 방법)
                if hasattr(shape, "text") and shape.text.strip():
                    text = shape.text.strip()
                    text_items.append({
                        'shape_index': shape_idx,
                        'original_text': text,
                        'shape_type': type(shape).__name__
                    })
                    print(f"    텍스트 추출: '{text[:30]}...' (shape_idx={shape_idx})")
                
                # 2. 그룹화된 도형 처리
                if hasattr(shape, "shapes"):
                    for child_idx, child_shape in enumerate(shape.shapes):
                        if hasattr(child_shape, "text") and child_shape.text.strip():
                            text = child_shape.text.strip()
                            text_items.append({
                                'shape_index': shape_idx,
                                'child_idx': child_idx, 
                                'original_text': text,
                                'shape_type': 'GroupedShape',
                                'is_grouped': True
                            })
                            print(f"    그룹 내 텍스트 추출: '{text[:30]}...' (shape_idx={shape_idx}, child_idx={child_idx})")
                
                # 3. 테이블 처리
                if isinstance(shape, GraphicFrame) and hasattr(shape, "table"):
                    try:
                        table = shape.table
                        print(f"    테이블 발견 (shape_idx={shape_idx})")
                        
                        # 테이블의 모든 셀 처리
                        for row_idx, row in enumerate(table.rows):
                            for col_idx, cell in enumerate(row.cells):
                                if cell.text.strip():
                                    text = cell.text.strip()
                                    text_items.append({
                                        'shape_index': shape_idx,
                                        'original_text': text,
                                        'shape_type': 'TableCell',
                                        'is_table': True,
                                        'row_idx': row_idx,
                                        'col_idx': col_idx
                                    })
                                    print(f"      테이블 셀 텍스트 추출 [{row_idx},{col_idx}]: '{text[:30]}...'")
                    except Exception as e:
                        print(f"      테이블 처리 중 오류: {str(e)}")
                
                # 4. 차트 처리
                if isinstance(shape, GraphicFrame) and hasattr(shape, "chart"):
                    try:
                        chart = shape.chart
                        if hasattr(chart, "chart_title") and chart.chart_title and chart.chart_title.text_frame.text.strip():
                            text = chart.chart_title.text_frame.text.strip()
                            text_items.append({
                                'shape_index': shape_idx,
                                'original_text': text,
                                'shape_type': 'ChartTitle',
                                'is_chart': True
                            })
                            print(f"    차트 제목 추출: '{text[:30]}...' (shape_idx={shape_idx})")
                    except Exception as e:
                        print(f"      차트 처리 중 오류: {str(e)}")
                
                # 5. SmartArt 및 기타 다이어그램 처리
                if isinstance(shape, GraphicFrame) and hasattr(shape, "graphic"):
                    try:
                        if hasattr(shape.graphic, "graphic_data"):
                            # SmartArt에서 텍스트 찾기 시도
                            if hasattr(shape, "text") and shape.text.strip():
                                text = shape.text.strip()
                                text_items.append({
                                    'shape_index': shape_idx,
                                    'original_text': text,
                                    'shape_type': 'GraphicData',
                                    'is_graphic': True
                                })
                                print(f"    그래픽 데이터 텍스트 추출: '{text[:30]}...' (shape_idx={shape_idx})")
                    except Exception as e:
                        print(f"      그래픽 처리 중 오류: {str(e)}")
                
            except Exception as e:
                print(f"    도형 {shape_idx} 처리 중 오류: {str(e)}")
                            
        return text_items
    
    def translate_text(self, text: str, target_language: str, source_language: str = 'auto') -> str:
        """Amazon Bedrock Claude 3.5 Sonnet을 사용하여 텍스트를 번역합니다."""
        import time
        import botocore.exceptions
        
        language_names = {
            'ko': '한국어', 'en': '영어', 'ja': '일본어', 'zh': '중국어',
            'fr': '프랑스어', 'de': '독일어', 'es': '스페인어', 'it': '이탈리아어',
            'pt': '포르투갈어', 'ru': '러시아어'
        }
        
        target_lang_name = language_names.get(target_language, target_language)
        
        prompt = f"""다음 텍스트를 {target_lang_name}로 번역해주세요. 
번역할 때 다음 사항을 고려해주세요:
1. 원문의 의미와 뉘앙스를 정확히 전달
2. 자연스러운 표현 사용
3. 전문 용어는 해당 언어의 표준 용어 사용
4. 서식이나 특수문자는 그대로 유지
5. 번역된 텍스트만 출력 (설명이나 부가 정보 없이)

번역할 텍스트:
{text}

번역:"""

        max_retries = 5  # 최대 5회 재시도
        retry_delay = 5  # 초 단위 대기 시간
        
        for attempt in range(max_retries):
            try:
                body = {
                    "anthropic_version": "bedrock-2023-05-31",
                    "max_tokens": 4000,
                    "messages": [
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ],
                    "temperature": 0.0
                }
                
                response = self.bedrock_client.invoke_model(
                    modelId=self.model_id,
                    body=json.dumps(body)
                )
                
                response_body = json.loads(response['body'].read())
                translated_text = response_body['content'][0]['text'].strip()
                
                return translated_text
                
            except botocore.exceptions.ClientError as e:
                error_code = e.response.get('Error', {}).get('Code', '')
                
                # ThrottlingException 발생 시 재시도
                if error_code == 'ThrottlingException' or 'ThrottlingException' in str(e):
                    if attempt < max_retries - 1:  # 마지막 시도가 아닌 경우에만 재시도
                        wait_time = retry_delay * (attempt + 1)  # 점진적으로 대기 시간 증가
                        print(f"  API 제한으로 인한 오류 발생: {wait_time}초 후 재시도 ({attempt+1}/{max_retries})...")
                        time.sleep(wait_time)
                    else:
                        print(f"  최대 재시도 횟수 도달: {str(e)}")
                        return text  # 최대 재시도 횟수 도달 시 원본 텍스트 반환
                else:
                    print(f"  번역 중 오류 발생: {str(e)}")
                    return text  # 다른 예외 발생 시 원본 텍스트 반환
            
            except Exception as e:
                print(f"  번역 중 오류 발생: {str(e)}")
                return text  # 다른 예외 발생 시 원본 텍스트 반환
    
    def translate_presentation(self, input_file: str, output_file: str, target_language: str) -> bool:
        """PowerPoint 프레젠테이션을 번역합니다."""
        
        if target_language not in self.supported_languages:
            print(f"지원하지 않는 언어입니다. 지원 언어: {list(self.supported_languages.keys())}")
            return False
        
        try:
            # PowerPoint 파일 로드
            prs = Presentation(input_file)
            total_slides = len(prs.slides)
            
            print(f"번역 시작: {total_slides}개 슬라이드를 {self.supported_languages[target_language]}로 번역합니다...")
            
            # 성공/실패 슬라이드 카운트
            success_count = 0
            failed_count = 0
            
            # 각 슬라이드 처리
            for slide_idx, slide in enumerate(prs.slides):
                try:
                    print(f"슬라이드 {slide_idx + 1}/{total_slides} 처리 중...")
                    
                    # 텍스트 추출
                    text_items = self.extract_text_from_slide(slide)
                    
                    # 각 텍스트 요소 번역
                    for item in text_items:
                        try:
                            original_text = item['original_text']
                            
                            # 공백만 있는 텍스트는 건너뛰기
                            if not original_text.strip():
                                continue
                            
                            # 번역 수행
                            translated_text = self.translate_text(original_text, target_language)
                            
                            # 번역된 텍스트를 슬라이드에 적용
                            shape = slide.shapes[item['shape_index']]
                            
                            # 텍스트 번역 적용 (서식 보존 방식)
                            applied = False
                            
                            # 1. 테이블 셀인 경우 - 텍스트 프레임 활용
                            if item.get('is_table', False):
                                table = shape.table
                                row_idx = item['row_idx']
                                col_idx = item['col_idx']
                                try:
                                    # 텍스트 프레임을 통한 단락별 서식 보존 적용
                                    cell = table.cell(row_idx, col_idx)
                                    if hasattr(cell, "text_frame"):
                                        # 기존 단락 구조와 서식 유지
                                        text_frame = cell.text_frame
                                        
                                        # 텍스트프레임의 모든 단락 정보 백업
                                        paragraph_info = []
                                        for p in text_frame.paragraphs:
                                            run_info = []
                                            for r in p.runs:
                                                run_info.append({
                                                    'text': r.text,
                                                    'font': r.font,  # 폰트 객체 자체를 저장
                                                })
                                            paragraph_info.append({
                                                'runs': run_info,
                                                'alignment': p.alignment,
                                                'level': p.level
                                            })
                                        
                                        # 문단별로 나누어 처리
                                        translated_paras = translated_text.split('\n')
                                        
                                        # 원래 텍스트 지우기
                                        while len(text_frame.paragraphs) > 1:
                                            p = text_frame.paragraphs[-1]
                                            tr = p._element
                                            tr.getparent().remove(tr)
                                        
                                        # 첫 번째 단락만 남겨두고 텍스트 초기화
                                        if text_frame.paragraphs:
                                            first_p = text_frame.paragraphs[0]
                                            for run in first_p.runs:
                                                run.text = ""
                                            
                                            # 첫 번째 단락에 첫 번째 번역 텍스트 적용
                                            if translated_paras:
                                                # 단락에 실행이 있으면 첫 번째 실행 사용, 없으면 새 실행 추가
                                                if first_p.runs:
                                                    first_p.runs[0].text = translated_paras[0]
                                                else:
                                                    first_p.text = translated_paras[0]
                                            
                                            # 나머지 단락 추가
                                            for i, trans_para in enumerate(translated_paras[1:], 1):
                                                p = text_frame.add_paragraph()
                                                # 기존 서식 정보를 가능한 복구
                                                if i < len(paragraph_info):
                                                    p.alignment = paragraph_info[i]['alignment']
                                                    p.level = paragraph_info[i]['level']
                                                p.text = trans_para
                                        
                                        print(f"  테이블 셀 서식 보존 번역 완료 [{row_idx},{col_idx}]: '{original_text[:30]}...' -> '{translated_text[:30]}...'")
                                        applied = True
                                    else:
                                        # 텍스트 프레임이 없는 경우 직접 텍스트 설정
                                        cell.text = translated_text
                                        print(f"  테이블 셀 번역 완료 [{row_idx},{col_idx}]: '{original_text[:30]}...' -> '{translated_text[:30]}...'")
                                        applied = True
                                except Exception as e:
                                    print(f"  테이블 셀 번역 실패: {str(e)}")
                                    # 실패 시 간단한 방식으로 시도
                                    try:
                                        table.cell(row_idx, col_idx).text = translated_text
                                        print(f"  테이블 셀 기본 방식 번역 완료: {original_text[:30]}...")
                                        applied = True
                                    except Exception as e2:
                                        print(f"  테이블 셀 기본 방식도 실패: {str(e2)}")
                            
                            # 2. 차트 제목인 경우 - 서식 보존
                            if not applied and item.get('is_chart', False):
                                try:
                                    chart = shape.chart
                                    text_frame = chart.chart_title.text_frame
                                    
                                    # 서식 보존 방식으로 적용
                                    if text_frame.paragraphs:
                                        # 기존 단락의 서식 정보 저장
                                        para = text_frame.paragraphs[0]
                                        if para.runs:
                                            # 기존 폰트 정보 저장
                                            font = para.runs[0].font
                                            
                                            # 텍스트 지우기
                                            for run in para.runs:
                                                run.text = ""
                                                
                                            # 첫 번째 실행에 번역된 텍스트 적용
                                            para.runs[0].text = translated_text
                                        else:
                                            # 실행이 없으면 단락 텍스트 설정
                                            para.text = translated_text
                                    else:
                                        # 단락이 없으면 텍스트 프레임에 직접 설정
                                        text_frame.text = translated_text
                                    
                                    print(f"  차트 제목 서식 보존 번역 완료: '{original_text[:30]}...' -> '{translated_text[:30]}...'")
                                    applied = True
                                except Exception as e:
                                    print(f"  차트 제목 번역 실패: {str(e)}")
                                    # 실패 시 기본 방식 시도
                                    try:
                                        chart.chart_title.text_frame.text = translated_text
                                        applied = True
                                    except:
                                        pass
                            
                            # 3. 그룹화된 도형인 경우 - 서식 보존
                            if not applied and item.get('is_grouped', False):
                                try:
                                    child_idx = item['child_idx']
                                    child_shape = shape.shapes[child_idx]
                                    
                                    # 텍스트 프레임을 통한 서식 보존 적용
                                    if hasattr(child_shape, "text_frame"):
                                        text_frame = child_shape.text_frame
                                        
                                        # 문단별로 나누어 처리
                                        translated_paras = translated_text.split('\n')
                                        
                                        # 단락별 서식 정보 백업
                                        paragraph_info = []
                                        for p in text_frame.paragraphs:
                                            run_info = []
                                            for r in p.runs:
                                                run_info.append({
                                                    'text': r.text,
                                                    'font': r.font  # 폰트 객체 저장
                                                })
                                            paragraph_info.append({
                                                'runs': run_info,
                                                'alignment': p.alignment,
                                                'level': p.level
                                            })
                                        
                                        # 원래 텍스트 지우기 (첫 번째 단락 제외)
                                        while len(text_frame.paragraphs) > 1:
                                            p = text_frame.paragraphs[-1]
                                            tr = p._element
                                            tr.getparent().remove(tr)
                                        
                                        # 첫 번째 단락 처리
                                        if text_frame.paragraphs:
                                            first_p = text_frame.paragraphs[0]
                                            
                                            # 단락의 실행 처리
                                            if first_p.runs:
                                                # 기존 폰트 정보 적용
                                                run = first_p.runs[0]
                                                # 다른 실행 제거
                                                while len(first_p.runs) > 1:
                                                    r = first_p.runs[-1]
                                                    r._r.getparent().remove(r._r)
                                                
                                                # 첫 번째 실행에 텍스트 설정 
                                                if translated_paras:
                                                    run.text = translated_paras[0]
                                            else:
                                                if translated_paras:
                                                    first_p.text = translated_paras[0]
                                            
                                            # 나머지 단락 처리
                                            for i, para_text in enumerate(translated_paras[1:], 1):
                                                para = text_frame.add_paragraph()
                                                # 기존 서식 정보 복구
                                                if i < len(paragraph_info):
                                                    para.alignment = paragraph_info[i]['alignment']
                                                    para.level = paragraph_info[i]['level']
                                                para.text = para_text
                                    else:
                                        # 텍스트 프레임이 없는 경우 직접 텍스트 설정
                                        child_shape.text = translated_text
                                    
                                    print(f"  그룹 내 도형 서식 보존 번역 완료: '{original_text[:30]}...' -> '{translated_text[:30]}...'")
                                    applied = True
                                except Exception as e:
                                    print(f"  그룹 내 도형 번역 실패: {str(e)}")
                                    # 실패 시 기본 방식 시도
                                    try:
                                        child_shape.text = translated_text
                                        applied = True
                                    except:
                                        pass
                            
                            # 4. 일반 도형 - 텍스트 프레임 활용
                            if not applied:
                                try:
                                    if hasattr(shape, "text_frame"):
                                        text_frame = shape.text_frame
                                        
                                        # 문단별로 나누어 처리
                                        translated_paras = translated_text.split('\n')
                                        
                                        # 단락 서식 정보 백업
                                        paragraph_info = []
                                        for p in text_frame.paragraphs:
                                            run_info = []
                                            for r in p.runs:
                                                if hasattr(r, 'font'):
                                                    run_info.append({
                                                        'text': r.text,
                                                        'font': r.font
                                                    })
                                            paragraph_info.append({
                                                'runs': run_info,
                                                'alignment': p.alignment if hasattr(p, 'alignment') else None,
                                                'level': p.level if hasattr(p, 'level') else 0
                                            })
                                        
                                        # 원래 텍스트 지우기 (첫 번째 단락 제외)
                                        while len(text_frame.paragraphs) > 1:
                                            p = text_frame.paragraphs[-1]
                                            tr = p._element
                                            tr.getparent().remove(tr)
                                        
                                        # 첫 번째 단락 처리
                                        if text_frame.paragraphs:
                                            first_p = text_frame.paragraphs[0]
                                            
                                            # 실행 처리
                                            if first_p.runs:
                                                # 첫 번째를 제외한 모든 실행 제거
                                                while len(first_p.runs) > 1:
                                                    r = first_p.runs[-1]
                                                    r._r.getparent().remove(r._r)
                                                
                                                # 첫 번째 실행에 텍스트 설정
                                                if translated_paras:
                                                    first_p.runs[0].text = translated_paras[0]
                                            else:
                                                if translated_paras:
                                                    first_p.text = translated_paras[0]
                                            
                                            # 나머지 단락 추가
                                            for i, para_text in enumerate(translated_paras[1:], 1):
                                                para = text_frame.add_paragraph()
                                                # 기존 서식 정보 복구
                                                if i < len(paragraph_info) and paragraph_info[i]['alignment'] is not None:
                                                    para.alignment = paragraph_info[i]['alignment']
                                                if i < len(paragraph_info):
                                                    para.level = paragraph_info[i]['level']
                                                para.text = para_text
                                        
                                        print(f"  텍스트 프레임 서식 보존 번역 완료: '{original_text[:30]}...' -> '{translated_text[:30]}...'")
                                        applied = True
                                    elif hasattr(shape, "text"):
                                        # 텍스트 프레임이 없는 경우 기본 텍스트 속성 사용
                                        shape.text = translated_text
                                        print(f"  기본 텍스트 번역 완료: '{original_text[:30]}...' -> '{translated_text[:30]}...'")
                                        applied = True
                                except Exception as e:
                                    print(f"  텍스트 적용 실패: {str(e)}")
                                    
                            if not applied:
                                print(f"  적용 방법을 찾을 수 없음: {item.get('shape_type', '알 수 없음')}")
                        except Exception as e:
                            print(f"  텍스트 항목 번역 실패: {str(e)}")
                    
                    success_count += 1
                except Exception as e:
                    print(f"  슬라이드 {slide_idx + 1} 처리 중 오류 발생: {str(e)}")
                    failed_count += 1
            
            # 번역된 파일 저장
            try:
                prs.save(output_file)
                print(f"\n번역 완료! 저장된 파일: {output_file}")
                print(f"슬라이드 처리 결과: 성공 {success_count}개, 실패 {failed_count}개")
                return success_count > 0  # 하나 이상의 슬라이드가 성공적으로 처리되었으면 성공으로 간주
            except Exception as e:
                print(f"파일 저장 중 오류 발생: {str(e)}")
                return False
            
        except Exception as e:
            print(f"프레젠테이션 로드 중 오류 발생: {str(e)}")
            return False
    
    def show_supported_languages(self):
        """지원하는 언어 목록을 출력합니다."""
        print("지원하는 언어:")
        for code, name in self.supported_languages.items():
            print(f"  {code}: {name}")

def main():
    translator = PowerPointTranslator()
    
    print("=== PowerPoint 번역기 ===")
    print("Amazon Bedrock Claude 3.5 Sonnet 사용\n")
    
    # 지원 언어 출력
    translator.show_supported_languages()
    
    # 사용자 입력
    input_file = input("\n번역할 PowerPoint 파일 경로를 입력하세요: ").strip()
    
    if not os.path.exists(input_file):
        print("파일이 존재하지 않습니다.")
        return
    
    target_language = input("번역할 언어 코드를 입력하세요 (예: ko, en, ja): ").strip().lower()
    
    # 출력 파일명 생성
    base_name = os.path.splitext(input_file)[0]
    output_file = f"{base_name}_translated_{target_language}.pptx"
    
    # 번역 실행
    success = translator.translate_presentation(input_file, output_file, target_language)
    
    if success:
        print(f"\n✅ 번역이 성공적으로 완료되었습니다!")
        print(f"📁 번역된 파일: {output_file}")
    else:
        print("❌ 번역 중 오류가 발생했습니다.")

if __name__ == "__main__":
    main()
