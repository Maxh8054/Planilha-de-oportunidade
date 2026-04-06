---
Task ID: 4
Agent: fullstack-developer
Task: Create opportunity types, constants, and utils components

Work Log:
- Created src/components/opportunity/types.ts with OpportunityRecord interface, ExportType and EstoqueFilterType types
- Created src/components/opportunity/constants.ts with all app constants including SHEET_NAMES, COLORS, STATUS_CONFIG, COLUMN_MAP, and filter options
- Created src/components/opportunity/utils.ts with utility functions: parseDate, parseNumber, parseCriticidade, determineStatus, getStockInfo, formatDateBR

Stage Summary:
- All 3 opportunity business logic files created successfully
- Files are in src/components/opportunity/ directory
- No lint errors in created files

---
Task ID: 5
Agent: main
Task: Make filter sections collapsible (bloco suspenso) for both Overview and Opportunities tabs

Work Log:
- Imported Collapsible, CollapsibleContent, CollapsibleTrigger from shadcn/ui
- Added showOverviewFilters and showOppFilters state (default: collapsed/false)
- Wrapped Overview tab filter Card in Collapsible with click-to-toggle header
- Wrapped Opportunities tab filter Card in Collapsible with click-to-toggle header
- Added ChevronDown icon with rotate animation to indicate open/closed state
- Added "X ativos" badge count when filters are active
- "Limpar" (clear) button only appears when filters are active
- Hover effect on header for better UX feedback

Stage Summary:
- Both filter sections are now collapsible (default: collapsed)
- Header shows active filter count badge when filters are applied
- ChevronDown icon animates to indicate expanded/collapsed state
- No lint errors in modified code
