const getUserRole = (sheet, userId) => {
  const collaborator = sheet.collaborators.find(
    (item) => item.userId.toString() === userId.toString()
  );

  return collaborator ? collaborator.role : null;
};

const canRead = (role) => ["owner", "admin", "editor", "viewer"].includes(role);
const canEdit = (role) => ["owner", "admin", "editor"].includes(role);
const canManage = (role) => role === "owner";
const canBypassRowLocks = (role) => role === "owner" || role === "admin";
const canManageSheetUsers = (role) => ["owner", "admin"].includes(role);

const canAssignSheetRole = (actorRole, currentRole, nextRole) => {
  if (currentRole === "owner" || nextRole === "owner") return false;
  if (actorRole === "owner") return ["admin", "editor", "viewer"].includes(nextRole);
  if (actorRole === "admin") {
    return currentRole !== "admin" && ["editor", "viewer"].includes(nextRole);
  }
  return false;
};

const canRemoveSheetCollaborator = (actorRole, targetRole) => {
  if (targetRole === "owner") return false;
  if (actorRole === "owner") return true;
  return actorRole === "admin" && targetRole !== "admin";
};

const getWorkspaceRole = (workspace, userId) => {
  const member = workspace.members.find(
    (item) => item.userId.toString() === userId.toString()
  );

  return member ? member.role : null;
};

const canManageWorkspace = (role) => role === "admin";
const canUseWorkspace = (role) => ["admin", "member", "viewer"].includes(role);

module.exports = {
  canAssignSheetRole,
  canBypassRowLocks,
  canEdit,
  canManage,
  canManageSheetUsers,
  canManageWorkspace,
  canRead,
  canRemoveSheetCollaborator,
  canUseWorkspace,
  getUserRole,
  getWorkspaceRole,
};
