# Architecture Decision Records

## ADR-001: Comment Personalization

**Date:** 2026-01-20

**Status:** Implemented

### Context

The codebase had functional comments that described *what* code did, but lacked the *why* and design rationale. Comments were generic and impersonal, making it harder to understand the reasoning behind implementation choices.

### Decision

Rewrite approximately 60% of comments across all 9 VBA modules to be intent-focused and explain design rationale rather than just describing operations.

### Standards Applied

**Architecture Comments**
- Use first-person to indicate deliberate design: "I designed this to...", "My approach here..."
- Explain why a particular approach was chosen over alternatives

**Bug Fix Comments**
- Document fixes with context: "Fixed: ...", "Had to handle..."
- Explain what was broken and why the fix works

**Business Context Comments**
- Link code to business requirements: "Business requirement: ...", "Per workflow..."
- Help future maintainers understand domain constraints

**Technical Constraint Comments**
- Document platform limitations: "VBA quirk: ...", "Excel limitation..."
- Prevent others from "fixing" intentional workarounds

### What Was Preserved

- Section dividers (====== lines)
- Single-word structural labels
- Cell reference documentation (e.g., "C17", "H5:I12")
- Debug.Print statements and their messages

### Modules Updated

| Module | Lines Changed | Focus Areas |
|--------|---------------|-------------|
| EntryPoint.bas | ~155 | UI setup rationale, event handling design |
| TX_NewUsage.bas | ~82 | Template selection logic, form layout decisions |
| TX_Return.bas | ~109 | Bulk entry design, CRDB integration notes |
| TX_Swap.bas | ~79 | Equipment type workflow, Dealer ID handling |
| PathHelper.bas | ~46 | Security-focused sanitization rationale |
| Dispatcher.bas | ~23 | Order type routing design |
| FileHelper.bas | ~31 | Template validation approach |
| SharePointHelper.bas | ~27 | Registry lookup strategy |
| Config.bas | ~10 | Configuration organization |

### Consequences

**Positive:**
- Code now communicates intent, not just mechanics
- Easier onboarding for new maintainers
- Design decisions are preserved alongside code
- Workarounds for VBA/Excel quirks are documented

**Neutral:**
- No runtime behavior changes
- No impact on existing functionality

**Negative:**
- None identified
