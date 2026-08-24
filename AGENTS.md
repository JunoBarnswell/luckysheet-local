# Repository Development Constraints

## GitHub PR workflow

- All subsequent development must happen on a dedicated non-default branch using the `codex/` prefix.
- Do not commit directly to `main` and do not push feature work to `main`.
- Every completed change must be delivered through a GitHub Pull Request targeting the repository's default branch.
- Before opening a PR, include the implementation summary, affected contracts, verification evidence, known blockers, and rollback considerations in the PR description.
- Keep unrelated user changes intact; never use destructive reset or checkout commands to clean the worktree.
- A PR is not considered complete until its branch has passed the repository verification gates and the PR status/checks are recorded.

## Architecture and product constraints

- Implement the spreadsheet runtime as one canonical semantic chain across model, render, interaction, command, persistence, OOXML, and backend layers.
- Do not add compatibility bridges, aliases, shims, double-write paths, parallel read models, or UI-only repair logic.
- Destructive refactors are allowed when they remove obsolete abstractions and converge ownership; update all consumers and contracts in the same change.
- `PaneMap` is the only coordinate ownership authority. Frozen pane visible ranges must be mutually exclusive, and hit, selection, editing, and commit addresses must remain identical.
- Worksheet and Table AutoFilter state must have one resolved owner per overlapping range. All UI, commands, hidden-row projection, import, and export behavior must use the same typed filter resolution.
- Hidden rows or columns must never make canonical cell reads return `undefined`; visibility is a separate projection.
- Unknown OOXML parts, nodes, extensions, non-standard paths, and macro content must be preserved unless an explicit format conversion owns their removal.
- Do not use rendering deduplication, row truncation, hardcoded offsets, mock data, placeholder handlers, or silent degradation to conceal a broken upstream contract.

## Verification and acceptance

- Run architecture-first reasoning before verification; tests and builds do not replace semantic review.
- Required frontend and backend checks must be recorded in the PR when relevant.
- Spreadsheet UI changes require real in-app browser interaction, including console and network inspection. A build or unit test alone is insufficient for UI completion.
- Native Excel interoperability must be verified with the applicable real-file corpus. If desktop Excel is unavailable, mark that acceptance item as `Blocked` rather than claiming completion.

## Development execution constraints

- Prefer destructive remediation when the current abstraction has the wrong ownership or semantics; remove the obsolete path and update every consumer and contract in the same change.
- Do not implement patches, compatibility bridges, aliases, shims, parallel state paths, duplicate command chains, UI-only repairs, mock behavior, placeholder handlers, or silent fallbacks.
- Before editing, reason through the product design, canonical data flow, ownership boundaries, state transitions, permission behavior, persistence semantics, and failure modes. Locate the real entrypoint, caller chain, state source, data contract, and verification path first.
- Treat the model, render, interaction, command, persistence, OOXML, collaboration, and backend layers as one semantic chain. Every visible product action must resolve to that chain and have real behavior.
- Complete the agreed implementation pass before running frequent compile or test cycles. Prefer one coherent development pass followed by the full verification gates; use intermediate checks only when a dependency error would otherwise block continued implementation.
- Do not declare completion from a green build alone. Completion requires architecture review, real behavior, persistence/interoperability evidence, and the applicable in-app browser acceptance.
