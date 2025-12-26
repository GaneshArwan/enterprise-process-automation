# ProcessPilot — Enterprise Automation & Workload Orchestration System

[![Google Apps Script](https://img.shields.io/badge/Built%20With-Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google)](https://developers.google.com/apps-script)
[![SAP Integration](https://img.shields.io/badge/Integration-SAP%20GUI%20Scripting-008FD3?style=for-the-badge&logo=sap)](https://www.sap.com)
[![Status](https://img.shields.io/badge/Status-Production%20Ready-success?style=for-the-badge)]()

## 🚀 Overview

**ProcessPilot** is an enterprise-grade automation system designed to orchestrate the lifecycle of high-volume business requests. Acting as a robust middleware between user input interfaces (Google Sheets/Web Forms) and enterprise ERP systems (SAP), it automates validation, approval workflows, workload distribution, and script generation.

This system was engineered to solve high-frequency concurrency issues and manual processing bottlenecks, successfully **reducing manual processing time by ~70%** and **eliminating data collision errors** through a custom distributed locking architecture.

## 🌟 Key Features

### 1. Advanced Concurrency Control (Distributed Locking)
Standard Google Apps Script locks are insufficient for high-frequency concurrent edits. This project implements a custom **Distributed Key-Lock System** (`src/utils/wrapper_utils.js`) that provides granular, key-based locks similar to database row-locking.
- **Mechanism:** Uses `LockService` combined with `CacheService` to create granular, key-based locks (e.g., locking a specific Request ID rather than the entire spreadsheet).
- **Resilience:** Implements exponential backoff retry logic (`BACKOFF_CFG`) to handle contention gracefully without failing transactions.

### 2. Intelligent Workload Orchestration
The system autonomously routes requests to agents using a **Scaled Round Robin** algorithm (`src/classes/request_allocator.js`).
- **Real-Time Balancing:** Evaluates agent availability (Active/On Leave) and current workload weight (estimated processing time in seconds).
- **Matrix & BAU Logic:** Supports complex routing matrices for specific request types while maintaining a fallback "Business As Usual" (BAU) logic for general requests.

### 3. High-Performance Caching Layer
To overcome Google Apps Script execution limits, the system features a sophisticated caching layer (`src/utils/master_config_utils.js`).
* **Configuration Caching:** Heavily caches configuration data (Approvers, SLAs, Baselines) to minimize expensive `SpreadsheetApp` read operations.
* **Smart Invalidation:** Automatically invalidates cache entries when updates occur, ensuring data consistency without sacrificing performance.

### 4. SAP Automation Integration (RPA)
Bridge the gap between data entry and ERP with the **Script Factory** (`src/classes/script.js`).
- **Dynamic Generation:** Parses validated spreadsheet data and dynamically generates SAP GUI VBScripts (e.g., for `ME11`, `PIR` creation).
- **Efficiency:** Enables agents to perform bulk uploads to SAP with a single click, reducing manual data entry errors to near zero.

### 5. Robust Object-Oriented Architecture
Built using modern JavaScript (ES6+) patterns to ensure maintainability and scalability.
* **Domain Models:** `Request`, `Activity`, and `Attachment` classes encapsulate core business logic.
* **Handlers:** Specialized handlers (`RequestHandler`, `ActivityHandler`) manage distinct lifecycle phases.
* **Validation Engine:** `AttachmentValidator` performs complex regex and dependency checks before data is processed.

## 📂 Project Structure

```text
root/
├── app.js                  # Application entry point (Triggers: onEdit, onSubmit, Time-driven)
├── api.js                  # REST API endpoints for external integrations
├── classes/                # Core Business Logic Models
│   ├── activity.js         # Manages row data state and CRUD operations
│   ├── archive.js          # Archival logic for historical data
│   ├── attachment.js       # Manages file attachments and drive permissions
│   ├── attachment_validator.js # Complex validation logic for input data
│   ├── email.js            # Email notification factory and templates
│   ├── request.js          # Main orchestrator for the request lifecycle
│   ├── request_allocator.js # Workload distribution algorithms
│   ├── request_logger.js   # Audit logging system
│   └── script.js           # SAP VBScript generation factory
├── configs/                # Configuration (Sanitized)
│   ├── constants.js        # System constants, IDs, and Environment variables
│   └── enums.js            # System enumerations, Status codes, Mapping objects
├── handlers/               # Workflow Orchestrators
│   ├── activity_handler.js # Handles data synchronization (Master <-> Child)
│   ├── attachment_handler.js # Manages template copying and folder creation
│   ├── migration_handler.js # Utilities for data migration and fixes
│   └── request_handler.js  # Main workflow logic (Approvals, Rejections, Routing)
└── utils/                  # Shared Utility Libraries
    ├── _date_utils.js      # Date formatting and calculation helpers
    ├── activity_utils.js   # Helpers for reading/parsing sheet activity rows
    ├── attachment_utils.js # Helpers for attachment spreadsheet operations
    ├── drive_utils.js      # Google Drive API wrappers
    ├── email_utils.js      # HTML email formatting and validation
    ├── enum_utils.js       # Enumeration parsers
    ├── logging_utils.js    # Centralized logging utilities
    ├── master_config_utils.js # Caching and configuration retrieval logic
    ├── menu_utils.js       # Custom UI menu functions for Google Sheets
    ├── request_utils.js    # Request ID generation and tracking
    ├── sheet_utils.js      # Low-level spreadsheet operations and optimization
    ├── string_utils.js     # String manipulation helpers
    ├── workload_manager.js # Agent workload property management
    └── wrapper_utils.js    # Distributed locking & concurrency mechanisms
```

## 🔁 Workflow (Preview & Full Docs)

Quick recommendation: keep a concise summary here and link to the detailed PDFs in the docs/ folder so recruiters and readers get an immediate overview but can open the full flowcharts if they want.

- Full diagrams and detailed flowcharts: [docs/01_Request_Ingestion_Pipeline.pdf](./docs/01_Request_Ingestion_Pipeline.pdf), [docs/02_Approval_Workflow_Engine.pdf](./docs/02_Approval_Workflow_Engine.pdf), [docs/03_Intelligent_Workload_Orchestration.pdf](./docs/03_Intelligent_Workload_Orchestration.pdf), [docs/04_Agent_Assisted_Execution_Workflow.pdf](./docs/04_Agent_Assisted_Execution_Workflow.pdf)


<summary>Compact workflow summary</summary>

1. Intake — submission via Sheets/Form; timestamp & initial notification.  
2. Validate & Lock — data checks; distributed key-lock to avoid collisions. ([Ingestion Pipeline](./docs/01_Request_Ingestion_Pipeline.pdf))  
3. Approval Loop — requester validation then approvers; handle send-back/reject/auto-approve. ([Approval Workflow](./docs/02_Approval_Workflow_Engine.pdf))  
4. Allocation — matrix or BAU; filter busy agents, balance by workload, round-robin tie-breaker. ([Workload Orchestration](./docs/03_Intelligent_Workload_Orchestration.pdf))  
5. Execution — agent claims task, processes or generates RPA script; updates child/master sheets. ([Agent-Assisted Execution](./docs/04_Agent_Assisted_Execution_Workflow.pdf))  
6. Close — logging, notifications, cache invalidation, and archival. ([Agent-Assisted Execution](./docs/04_Agent_Assisted_Execution_Workflow.pdf))


<details>
## System Architecture & Workflows

This system is an event-driven, high-concurrency orchestration platform built to ensure data integrity, predictable SLA behavior, and efficient human+RPA workflows.

### 01. Request Ingestion Pipeline
- **Objective:** Validates and processes high-volume incoming data streams.
- **Mechanism:** Uses an interval-driven trigger to batch-process new submissions, perform gatekeeping checks (Duplicate Detection, Data Integrity, Expiration) and construct a "Sync Context" for requestors and approvers before handing off to the approval engine.
- **Key Tech:** Event Throttling, Cache-Based Validation, Context Construction.

### 02. Approval Workflow Engine
- **Objective:** Orchestrates multi-level decision gates and automated compliance checks.
- **Mechanism:** Iterates through a hierarchy of contexts (Requester → Approvers), handling states like NEED REVIEW, AUTO-APPROVE, and SEND BACK. On final approval it computes SLA baselines, locks the record, and triggers the workload allocator.
- **Key Tech:** State Machine Logic, SLA Calculation, Dynamic Routing.

### 03. Intelligent Workload Orchestration
- **Objective:** Autonomously assigns tasks to agents based on real-time availability and load.
- **Mechanism:** Dual-strategy algorithm:
    - **Matrix Distribution:** Routes specific request types using a skills matrix.
    - **Load Balancing (BAU):** Filters available agents (ignoring "Busy"), then uses a Least-Connections (minimum total seconds) algorithm with a Round-Robin tie-breaker.
- **Key Tech:** Round Robin Tie-Breaker, Workload Weighting, Resource Optimization.

### 04. Agent-Assisted Execution Workflow
- **Objective:** Hybrid Human-in-the-Loop workflow combining manual oversight with RPA tools.
- **Mechanism:**
    - **Path A (Claiming):** Grants dynamic Drive permissions and starts the SLA timer when an agent claims a task.
    - **Path B (Status Management):** Enforces validation gates (e.g., require timestamps before completion) and handles rejection loops.
    - **Path C (RPA Trigger):** Detects specific column edits to invoke the ScriptFactory, generating SAP VBScript files automatically for agent execution.
- **Key Tech:** Human-in-the-Loop (HITL), Dynamic ACL (Access Control Lists), Automated Script Generation.

For detailed diagrams and process flows, see the PDFs in the `docs/` folder referenced above.
</details>