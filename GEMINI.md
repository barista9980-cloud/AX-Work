# 🚨 CRITICAL WORKSPACE SECURITY & OPERATIONAL RULES

1. 📄 DEEP AI/LLM PDF CONTENT VERIFICATION (PDF 수료증 본문 AI/LLM 정밀 검수 의무):
   - NEVER rely solely on file titles or simple OCR/text pattern matching when inspecting certificate PDFs!
   - Even if the file title is missing, wrong, or contains typos (e.g. '성희록', 'Scan_001.pdf'), ALWAYS inspect the actual inner content using deep AI/LLM parsing to accurately extract the completion name, course title, completion date, and issuing authority.
   - Categorize and update the master register strictly based on the authoritative inner PDF text/content verification.

2. 🔒 MANDATORY 2-STEP CONFIRMATION FOR ALL EMAIL DISPATCH (메일 2차 재확인 필수의무):
   - EVEN IF the user tells you to send an email, you MUST ALWAYS pause and double-check with the user one more time before actually sending it!
   - Step 1: User says "send email".
   - Step 2: Show exact recipients, subject, and body, and ask: "전송 전 2차 재확인 요청: 이 메일을 실제로 발송할까요?"
   - Step 3: ONLY send after user gives explicit secondary confirmation ("네, 발송하세요").

3. 🛑 GIT PUSH PRE-APPROVAL MANDATE (깃허브 푸시 사전 승인 필수):
   - NEVER execute `git commit` or `git push` without explicit user permission beforehand.

4. 📊 MULTI-TENANT ASSET & STATUTORY TRAINING STANDARD:
   - Master Framework name: `폭스패밀리 (FoxFamily)`
   - Master Register: Single Source of Truth `2026년도_법정의무교육_통합수료현황대장_폭스패밀리.xlsx` located exclusively inside `2026년도_법정의무교육_이력` folder.
   - 9 Standard Columns: [순번, 성명, 부서, 소속법인, 성희롱예방(PDF), 장애인인식(PDF), 개인정보보호(PDF), 최종이수상태, 비고].
