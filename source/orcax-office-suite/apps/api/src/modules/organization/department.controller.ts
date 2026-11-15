import {
  Body,
  Controller,
  Delete,
  Get,
  HttpCode,
  Param,
  ParseUUIDPipe,
  Post,
} from '@nestjs/common';
import { CurrentAccess } from '../../common/http/request-context';
import { AccessContext } from '../../tenancy/access-context';
import { RequirePermission } from '../authorization/require-permission.decorator';
import { AssignMemberDto, CreateDepartmentDto, CreateGroupDto } from './department.dto';
import { DepartmentService } from './department.service';

@Controller()
export class DepartmentController {
  constructor(private readonly departments: DepartmentService) {}

  @Get('departments')
  @RequirePermission('workspace.department.manage')
  listDepartments(@CurrentAccess() access: AccessContext) {
    return this.departments.listDepartments(access);
  }

  @Post('departments')
  @HttpCode(201)
  @RequirePermission('workspace.department.manage')
  createDepartment(@CurrentAccess() access: AccessContext, @Body() body: CreateDepartmentDto) {
    return this.departments.createDepartment(access, body);
  }

  @Post('departments/:id/members')
  @HttpCode(200)
  @RequirePermission('workspace.department.manage')
  assignDepartment(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: AssignMemberDto,
  ) {
    return this.departments.assignDepartmentMember(access, id, body.userId);
  }

  @Delete('departments/:id/members/:userId')
  @RequirePermission('workspace.department.manage')
  removeDepartment(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Param('userId', ParseUUIDPipe) userId: string,
  ) {
    return this.departments.removeDepartmentMember(access, id, userId);
  }

  @Get('groups')
  @RequirePermission('workspace.department.manage')
  listGroups(@CurrentAccess() access: AccessContext) {
    return this.departments.listGroups(access);
  }

  @Post('groups')
  @HttpCode(201)
  @RequirePermission('workspace.department.manage')
  createGroup(@CurrentAccess() access: AccessContext, @Body() body: CreateGroupDto) {
    return this.departments.createGroup(access, body);
  }

  @Post('groups/:id/members')
  @HttpCode(200)
  @RequirePermission('workspace.department.manage')
  assignGroup(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Body() body: AssignMemberDto,
  ) {
    return this.departments.assignGroupMember(access, id, body.userId);
  }

  @Delete('groups/:id/members/:userId')
  @RequirePermission('workspace.department.manage')
  removeGroup(
    @CurrentAccess() access: AccessContext,
    @Param('id', ParseUUIDPipe) id: string,
    @Param('userId', ParseUUIDPipe) userId: string,
  ) {
    return this.departments.removeGroupMember(access, id, userId);
  }
}
