import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { AccessContext } from '../../tenancy/access-context';
import { requireWorkspaceId } from '../../tenancy/workspace-scope';
import { AuthorizationService } from '../authorization/authorization.service';
import { can } from '../authorization/catalog';
import { License } from '../licensing/license.entity';
import { Membership } from '../organization/membership.entity';
import { Profile } from '../profile/profile.entity';
import { Template } from '../template/template.entity';

@Injectable()
export class AdminOverviewService {
  constructor(
    @InjectRepository(Membership)
    private readonly memberships: Repository<Membership>,
    @InjectRepository(License)
    private readonly licenses: Repository<License>,
    @InjectRepository(Profile)
    private readonly profiles: Repository<Profile>,
    @InjectRepository(Template)
    private readonly templates: Repository<Template>,
    private readonly authorization: AuthorizationService,
  ) {}

  async overview(actor: AccessContext) {
    this.authorization.assertPolicy(actor, { permission: 'workspace.read' });
    const workspaceId = requireWorkspaceId(actor);
    const widgets = [];

    if (can(actor.roleCode, 'workspace.read')) {
      const activeMembers = await this.memberships.count({
        where: { workspaceId, status: 'active' },
      });
      const suspendedMembers = await this.memberships.count({
        where: { workspaceId, status: 'suspended' },
      });
      widgets.push({
        id: 'members',
        title: 'Thành viên',
        href: '/thanh-vien',
        metrics: [
          { label: 'Đang hoạt động', value: activeMembers },
          { label: 'Tạm ngưng', value: suspendedMembers },
        ],
      });
    }

    if (can(actor.roleCode, 'license.read')) {
      const licenseCount = await this.licenses.count({ where: { workspaceId } });
      widgets.push({
        id: 'licenses',
        title: 'Giấy phép',
        href: '/giay-phep',
        metrics: [{ label: 'Bản ghi', value: licenseCount }],
      });
    }

    if (can(actor.roleCode, 'profile.read')) {
      const profileCount = await this.profiles.count({ where: { workspaceId } });
      const approved = await this.profiles.count({
        where: { workspaceId, status: 'approved' },
      });
      widgets.push({
        id: 'profiles',
        title: 'Hồ sơ',
        href: '/ho-so',
        metrics: [
          { label: 'Trong workspace', value: profileCount },
          { label: 'Đã duyệt', value: approved },
        ],
      });
    }

    if (can(actor.roleCode, 'template.read')) {
      const templateCount = await this.templates.count({ where: { workspaceId } });
      widgets.push({
        id: 'templates',
        title: 'Mẫu',
        href: '/mau',
        metrics: [{ label: 'Trong workspace', value: templateCount }],
      });
    }

    if (can(actor.roleCode, 'workspace.department.manage')) {
      widgets.push({
        id: 'departments',
        title: 'Phòng ban & nhóm',
        href: '/phong-ban',
        metrics: [{ label: 'Quản trị', value: 1 }],
      });
    }

    if (can(actor.roleCode, 'workspace.config.read')) {
      widgets.push({
        id: 'config',
        title: 'Cấu hình',
        href: '/cau-hinh',
        metrics: [{ label: 'Workspace', value: 1 }],
      });
    }

    if (can(actor.roleCode, 'workspace.audit.read')) {
      widgets.push({
        id: 'audit',
        title: 'Nhật ký',
        href: '/nhat-ky',
        metrics: [{ label: 'Tra cứu', value: 1 }],
      });
    }

    return {
      workspace: {
        id: workspaceId,
        kind: actor.workspaceKind,
        name: actor.workspaceName,
        role: actor.roleCode,
      },
      widgets,
    };
  }
}
