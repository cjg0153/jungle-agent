# Jungle Smart - Phase B

Version: 0.1.0

## Philosophy

Jungle does not try to out-think Claude or Codex.

Jungle provides:

1. Minimal routing
2. Maximum model autonomy
3. Source protection
4. Artifact integrity
5. Independent verification
6. Hard Gate
7. Cheapest-first validation

## Phase B principle

AI decides how to solve the task.

Jungle decides:

- which specialist should receive the task
- whether the source is protected
- whether the artifact exists
- whether deterministic checks pass
- whether independent QA is required
- whether the final result may pass the Hard Gate

## Canonical PPT SAFE engine

Jungle v3.2 SAFE Production Stable

The Phase B runtime must not modify the frozen stable candidate.

## Feature flags

Initially enabled:

- router

Initially disabled:

- execution
- deterministic_g1
- changed_scope
- watchdog
- blind_pairwise_qa
- reflow_v1
- worker_protocol

Features are activated only after independent smoke/regression testing.