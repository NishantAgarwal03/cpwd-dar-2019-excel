# Subject Tracker

This skill provides an automated, robust instruction-tracking system for multi-file projects. 

This skill MUST run automatically on EVERY user interaction without being asked (NO REMINDERS NEEDED).

## Goal

To maintain a historical record of all instructions, refinements, and scores for every active "Subject" (file, component, or feature) in a project.

## Automation Logic (ZERO MISS MODE)

### 0. Pre-Response Execution (MANDATORY)

Before generating ANY response:

1. Identify subject
2. Ensure subject_tracker.md exists
3. Update instruction log

ONLY AFTER completing above, generate response.

### 1. Subject Identification

- The "Subject" is defined as the file currently being edited, or the file name mentioned in the user's request (e.g., `organizer.py`).
- If no file is open or mentioned, use the specific feature or component name as the subject.
- Every detected Subject becomes automatically locked
- All locked subjects MUST be checked on every interaction.

### 2. Synchronization (Local Copy)

- Every time you interact with the tracker in a project, check if the project has a local copy of this skill at `.agents/skills/subject_tracker/SKILL.md`.
- If the local copy is missing, outdated, or if this is the first interaction, copy this Global `SKILL.md` definition to that local project path.
- This ensures the tracking logic "travels" with the project and is readable by other agents.

### 3. File Initialization

- MUST ALWAYS ensure subject_tracker.md file exists before proceeding. CREATE immediately if missing (NO USER PROMPT).
- When writing/modifying/creating project plan ensure subject_tracker.md` file exists in the project root.
- At the start of any new project interaction, ensure a `subject_tracker.md` file exists in the project root.
- If it does not exist, create it with a standard header.

### 4. Subject Sections

Each subject MUST have its own independent section in `subject_tracker.md` with:

- **Status**: `🔴 Active` or `🟢 Finalized`
- **Initial Score**: A self-assessment score (out of 10) for the work/state of the subject at the start.
- **Final Score**: The score (out of 10) provided when the subject is closed.
- **Satisfaction Level**: A record of the user's explicit feedback.
- **Instruction Log Table**: A chronological table of:
  - Timestamp (YYYY-MM-DD HH:MM)
  - Instruction (The atomic request from the user)
  - Status (In-Progress, Completed, or Finalized)

### 4. Trigger Conditions (MANDATORY EXECUTION)

- This skill is ALWAYS ACTIVE.
- **Finalization**: If the user says "I am happy," "it's okay," "looks good," or "we can close this," update the status to `🟢 Finalized`.
- **Scoring**: When the user requests a rating (e.g., "rate this out of 10"), provide a score and record it as the `Final Score` if closing.
- **Satisfaction Prompt**: Upon finalization, ASK the user: *"Tell me about your satisfaction with the modifications in [Subject]."*
- **Feedback Logging**: The user's response to the satisfaction prompt must be recorded in the `Satisfaction Level` field of that subject.

### 5. Independent Life-cycles

- Ensure multiple subjects can coexist in the same `subject_tracker.md` file without instructions bleeding between them.
- Always append new instructions under the correct `## Subject: [Name]` section.

## File Template (subject_tracker.md)

```markdown
# Project Instruction Tracker

## Subject: [Subject Name]
- **Status**: 🔴 Active
- **Initial Score**: X/10 (Quantitative self-assessment by LLM)
- **Final Score**: TBD
- **Satisfaction Level**: TBD (user feedback)

### Remarks
- [Summary of changes and results]

| Timestamp | Instruction | Status |
| :--- | :--- | :--- |
| YYYY-MM-DD HH:MM | Clear atomic request | In-Progress |
```

### **Section 6: Deep Archival (Maintenance Mandate)**

- **Trigger**: The agent SHOULD initiate archival when `subject_tracker.md` exceeds **200 lines** OR more than **5 subjects** are marked as `🟢 Finalized`.

- **Action**:
  
  1. Create (if missing) `.agents/skills/subject_tracker/ARCHIVE.md`.
  2. Move the entire block of any `🟢 Finalized` subjects (including their Instruction Logs) from the main tracker to the top of the `ARCHIVE.md` file.
  3. Replace the moved block in the main tracker `subject_tracker.md` with a single-line summary in an **"Archived Subjects"** table at the bottom of the file.

- **Archive Entry Format**:
  
  | Subject           | Final Score | Date Archived |
  | ----------------- | ----------- | ------------- |
  | `math_library.py` | 9.3/10      | 2024-04-09    |
