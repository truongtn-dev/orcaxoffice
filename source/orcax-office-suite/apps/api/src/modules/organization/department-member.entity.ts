import { Column, Entity, PrimaryColumn } from 'typeorm';

@Entity({ name: 'department_members' })
export class DepartmentMember {
  @PrimaryColumn({ name: 'department_id', type: 'uuid' })
  departmentId!: string;

  @PrimaryColumn({ name: 'user_id', type: 'uuid' })
  userId!: string;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;
}
