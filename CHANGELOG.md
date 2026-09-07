# Changelog

All notable changes to this project will be documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/), and this project adheres to [Semantic Versioning](https://semver.org/).

## [2.0.0] - 2026-08-21

### Added

- Added new datasource `Exchange-On-Premises-Get-GroupMembers` to retrieve current members of a distribution group
- Added comprehensive error handling with try/catch/finally blocks for all operations
- Added TLS 1.2 enforcement for secure connections
- Added explicit command imports to minimize memory usage and improve performance
- Added detailed audit logging with structured log objects for all operations
- Added duplicate member detection to skip adding users who are already members
- Added proper session cleanup in finally blocks to prevent session leaks
- Added inline documentation and comments for better code maintainability

### Changed

- Renamed all datasources with more descriptive names including "exchange-on-premises-distribution-group-manage-memberships" prefix
- Renamed datasource from `Exchange-distributiongroup-generate-table-wildcard` to `Exchange-On-Premises-Get-DistributionGroup-Wildcard-Name-Description-Mail`
- Renamed datasource from `Exchange-distributiongroup-generate-table-members` to `Exchange-On-Premises-Get-GroupMembers`
- Renamed datasource from `Exchange-user-generate-table-distributiongroups-manage-memberships` to `Exchange-On-Premises-Get-All-Users`
- Renamed task from `Exchange on-premise - Manage memberships distribution group` to `Exchange On-Premises - Distribution Group - Manage memberships`
- Improved connection establishment with better structured session parameters using splatting
- Enhanced error messages to include line numbers and specific error context
- Improved filter syntax for querying distribution groups with more flexible matching
- Refactored credential creation for better consistency across all scripts
- Updated session option configuration with explicit parameter definitions
- Improved logging messages with actionable context using `$actionMessage` variable

### Fixed

- Fixed session cleanup to properly dispose of Exchange PowerShell sessions
- Fixed error handling to properly catch and report remote exceptions
- Fixed member addition to handle "already present in collection" errors gracefully

## [1.0.2] - 2022-08-24

### Added

- Added version number and updated code for SA-agent and auditlogging

## [1.0.1] - 2021-11-16

### Added

- Added version number and updated all-in-one script

## [1.0.0] - 2021-04-29

Initial release of HelloID-Conn-SA-Full-Exchange-On-Premises-Distribution-Group-Manage-Memberships.

### Added

- Initial release for managing Exchange On-Premises Distribution Group memberships
- Support for searching distribution groups
- Support for adding members to distribution groups
- Support for removing members from distribution groups
- Basic error handling and audit logging
