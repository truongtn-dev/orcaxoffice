import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { In, Repository } from 'typeorm';
import { User } from './user.entity';

export interface UserSummary {
  id: string;
  email: string;
  displayName: string;
}

@Injectable()
export class UserDirectory {
  constructor(
    @InjectRepository(User)
    private readonly users: Repository<User>,
  ) {}

  async summaries(ids: readonly string[]): Promise<Map<string, UserSummary>> {
    if (ids.length === 0) {
      return new Map();
    }
    const rows = await this.users.find({
      where: { id: In([...ids]) },
      select: { id: true, email: true, displayName: true },
    });
    return new Map(rows.map((row) => [row.id, row]));
  }
}
