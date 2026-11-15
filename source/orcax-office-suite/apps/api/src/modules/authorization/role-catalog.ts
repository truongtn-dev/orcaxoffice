import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { EntityManager, Repository } from 'typeorm';
import { Role } from './role.entity';

@Injectable()
export class RoleCatalog {
  constructor(
    @InjectRepository(Role)
    private readonly roles: Repository<Role>,
  ) {}

  async requireByCode(code: string, manager?: EntityManager): Promise<Role> {
    const role = await this.repo(manager).findOne({ where: { code } });
    if (!role) {
      throw new Error(`Role ${code} is not seeded`);
    }
    return role;
  }

  async requireById(id: string, manager?: EntityManager): Promise<Role> {
    const role = await this.repo(manager).findOne({ where: { id } });
    if (!role) {
      throw new Error('Role is missing');
    }
    return role;
  }

  private repo(manager?: EntityManager): Repository<Role> {
    return manager ? manager.getRepository(Role) : this.roles;
  }
}
