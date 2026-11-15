import { Column, Entity, PrimaryColumn } from 'typeorm';

@Entity({ name: 'group_members' })
export class GroupMember {
  @PrimaryColumn({ name: 'group_id', type: 'uuid' })
  groupId!: string;

  @PrimaryColumn({ name: 'user_id', type: 'uuid' })
  userId!: string;

  @Column({ name: 'created_at', type: 'timestamptz' })
  createdAt!: Date;
}
