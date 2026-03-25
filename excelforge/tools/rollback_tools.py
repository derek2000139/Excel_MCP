# Rollback tools have been merged into snapshot tools
# - snapshot.list = rollback.list_snapshots + snapshot.get_stats
# - snapshot.preview = rollback.preview_snapshot
# - snapshot.manage (action: restore_snapshot) = rollback.restore_snapshot
# - snapshot.manage (action: restore_backup) = backup.restore_file
# - snapshot.delete (action: cleanup) = snapshot.run_cleanup

# This file is kept for reference but no longer registers any tools