export function getCreatedAtUpdatedAtValues(
  id: string
): { createdAt: number; updatedAt: number } {
  const file = DriveApp.getFileById(id);
  const createdAt = file.getDateCreated().getTime();
  const updatedAt = file.getLastUpdated().getTime();
  return { createdAt, updatedAt };
}
